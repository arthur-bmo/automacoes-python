import requests
import json
import pandas as pd
import time


def baixa_planilha_clientes(path):
    # Correct URL for listing clients
    url = "https://app.omie.com.br/api/v1/geral/clientes/"

    app_key = "app_key"
    app_secret = "app_secret"

    headers = {
        "Content-Type": "application/json"
    }

    all_clients = []
    unique_clients = {}
    page = 1
    total_pages = 1

    while page <= total_pages:
        payload = {
            "call": "ListarClientesResumido",
            "app_key": app_key,
            "app_secret": app_secret,
            "param": [
                {
                    "pagina": page,
                    "registros_por_pagina": 100,
                    "apenas_importado_api": "N"
                }
            ]
        }

        response = requests.post(url, headers=headers, data=json.dumps(payload))
        data_clientes = response.json()

        # Debug print to check the response structure
        print(f"Processing page {page} of {total_pages}")

        # Update the total_pages based on the response
        total_pages = data_clientes.get("total_de_paginas", 1)

        clientes_cadastro_resumido = data_clientes.get('clientes_cadastro_resumido')
        if clientes_cadastro_resumido is None:
            print(f"No 'clientes_cadastro_resumido' key found on page {page}")
            continue

        for client in clientes_cadastro_resumido:
            if client['cnpj_cpf'] not in unique_clients:
                unique_clients[client['cnpj_cpf']] = client
                all_clients.append(client)

        page += 1
        time.sleep(0.4)  # Add a delay between requests to avoid overwhelming the server

    # Convert the unique clients dictionary to a list
    all_clients = list(unique_clients.values())

    # Create a DataFrame from the list of clients
    df = pd.DataFrame(all_clients)  

    download_path = rf'{path}\lista_clientes.xlsx'

    # Save the DataFrame to an Excel file
    df.to_excel(download_path, index=False)

    print("Client list saved to lista_clientes.xlsx")

    return download_path