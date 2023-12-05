import pandas as pd
import re

# Carregar o arquivo Excel
df = pd.read_excel("base.xlsx", sheet_name="Page 1")

# Função para extrair o CNPJ
def extrair_cnpj(descricao):
    # Encontrar a posição do termo "CNPJ"
    pos_cnpj = descricao.find("CNPJ")
    
    # Se o termo "CNPJ" for encontrado
    if pos_cnpj != -1:
        # Encontrar o primeiro caractere numérico após "CNPJ"
        start_pos = pos_cnpj + len("CNPJ")
        while start_pos < len(descricao) and not descricao[start_pos].isdigit():
            start_pos += 1
        
        # Encontrar o próximo espaço em branco ou vazio
        end_pos = start_pos
        while end_pos < len(descricao) and descricao[end_pos] not in [' ', '']:
            end_pos += 1
        
        # Extrair o CNPJ
        cnpj = descricao[start_pos:end_pos]
        return cnpj.strip()
    
    return None

# Função para extrair a Razão Social
def extrair_razao_social(descricao):
    padrao = r'Razão Social\s*:\s*([\s\S]*?)(?=\s*Número do Pedido|$)'
    resultado = re.search(padrao, str(descricao), re.DOTALL)
    if resultado:
        return resultado.group(1).strip()
    return None

# Função para extrair o Número do Pedido
def extrair_numero_pedido(descricao):
    padrao = r'Número do Pedido\D*(\d{10})'
    resultado = re.search(padrao, descricao)
    if resultado:
        return resultado.group(1)
    return None

# Função para extrair a Nota Fiscal
def extrair_nota_fiscal(descricao):
    padrao = r'Nota Fiscal\D*(\d+)[^\d]*Valor da Nota Fiscal'
    resultado = re.search(padrao, descricao)
    if resultado:
        return resultado.group(1).strip()
    return None

# Função para extrair o Valor da Nota Fiscal
def extrair_valor_nf(descricao):
    padrao = r'Valor da Nota Fiscal\D*(\d[\d\.,]+)[^\d]*Data de Emissão NF'
    resultado = re.search(padrao, descricao)
    if resultado:
        valor_nf = resultado.group(1).strip()
        valor_nf = re.sub(r'[^\d\.,]', '', valor_nf)  # Extrair apenas dígitos, ".", e ","
        return valor_nf
    return None

# Função para extrair a Data de Emissão NF
def extrair_data_emissao_nf(descricao):
    padrao = r'Data de Emissão NF\D*([\d-]+)[^\d]*Despesa'
    resultado = re.search(padrao, descricao)
    if resultado:
        return resultado.group(1).strip()
    return None

# Função para extrair a Despesa
def extrair_tipo_despesa(descricao):
    padrao = r'Despesa\s*:\s*([\s\S]*?)(?=\s|Descrição|$)'
    resultado = re.search(padrao, str(descricao), re.DOTALL)
    if resultado:
        return resultado.group(1).strip()
    return None

# Aplicar as funções às respectivas colunas para criar as novas colunas
df['Cnpj'] = df['Descrição'].apply(extrair_cnpj)
df['Razão Social'] = df['Descrição'].apply(extrair_razao_social)
df['Numero de pedido'] = df['Descrição'].apply(extrair_numero_pedido)
df['Nota Fiscal'] = df['Descrição'].apply(extrair_nota_fiscal)
df['Valor'] = df['Descrição'].apply(extrair_valor_nf)
df['Data de Emissão NF'] = df['Descrição'].apply(extrair_data_emissao_nf)
df['Tipo de Despesa'] = df['Descrição'].apply(extrair_tipo_despesa)

# Salvar o DataFrame resultante em um novo arquivo Excel
df.to_excel("base_completa.xlsx", index=False)
