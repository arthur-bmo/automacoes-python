import requests
import json
import re

def encontra_codigo_cliente(favorecido, cpfpj, lista_clientes):
    clientes = lista_clientes['nome_fantasia'].tolist()
    cpfs = lista_clientes['cnpj_cpf'].tolist()
    codigo_omie = lista_clientes['codigo_cliente'].tolist()
    codigo_cliente_found = None
    for cpf in cpfs:
        if str(cpf).replace(".", "").replace("-", "").replace("/",'') == str(cpfpj).replace(".", "").replace("-", "").replace("/",''):
            index = cpfs.index(cpf)
            codigo_cliente_found = True
            return codigo_omie[index]
    if not codigo_cliente_found:
        print("#### CLIENTE NÃO ENCONTRADO NO OMIE ####")

def registra_cliente_omie(app_key, app_secret, favorecido, cpfpj):
    url = "https://app.omie.com.br/api/v1/geral/clientes/"
    payload = {
        "call": "IncluirCliente",
        "app_key": app_key,
        "app_secret": app_secret,
        "param": [
            {
                "codigo_cliente_integracao": str(cpfpj).replace(".","").replace("/","").replace("-",""),  # Usa CPF/CNPJ como identificador único
                "razao_social": favorecido,
                "cnpj_cpf": cpfpj,
                "nome_fantasia": favorecido,
            }
        ]
    }

    headers = {"Content-Type": "application/json"}
    response = requests.post(url, headers=headers, data=json.dumps(payload))
    data_response = response.json() 

    if "codigo_cliente_omie" in data_response:
        print(data_response)
        return data_response["codigo_cliente_omie"]
    else:
        match = re.search(r'nCod \[(\d+)\]', data_response['faultstring'])
        if match:
            ncod = match.group(1)
            print(f"Cliente já cadastrado com nCod: {ncod}")
            return ncod
        print("Erro ao cadastrar cliente:", data_response)
        return None