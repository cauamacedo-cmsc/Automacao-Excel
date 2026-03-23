# importar a biblioteca pandas e dar o apelido "pd"
# pandas é usado para ler e manipular planilhas Excel
import pandas as pd

tabela="MODELO_LOTE - KITDOC V2.0.xlsx"

# lê especificamente a aba RUBRICA
rubrica_id=pd.read_excel("abreviacoes.xlsx",sheet_name="RUBRICA")

rubricas={
    "CCC":5,
    "CCB":6,
    "CHP":8,
    "CBC":9}

preencher=pd.read_excel("MODELO_LOTE - KITDOC V2.0.xlsx")

#converte coluna "rubrica" para object, assim aceita número e texto
preencher["rubrica"]=preencher["rubrica"].astype(object)

# iterrows percorre linha por linha da planilha
# linha = dados daquela linha
for i,linha in preencher.iterrows():
    nome=linha["descricao"]

    #ignora linhas vazias
    if pd.isna(nome):
        preencher.loc[i,"rubrica"]="#ERRO#"
        continue

    #split separa o testo em palavras
    palavras=str(nome).split()

    #pega a ultima palavra do texto
    sigla_rubrica=palavras[-1]

    # get procura a sigla dentro do dicionário
    rubrica_encontrada=rubricas.get(sigla_rubrica)

    if rubrica_encontrada:
        preencher.loc[i,"rubrica"]=rubrica_encontrada

    else:
        preencher.loc[i,"rubrica"]="#ERRO#"

# salva a planilha já preenchida
preencher.to_excel("MODELO_LOTE - PREENCHIDO.xlsx", index=False)