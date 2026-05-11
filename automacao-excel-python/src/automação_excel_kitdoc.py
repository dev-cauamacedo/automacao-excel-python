# importar a biblioteca pandas e dar o apelido "pd"
# pandas é usado para ler e manipular planilhas Excel
import pandas as pd

# abrir planilha
preencher=pd.read_excel("MODELO_LOTE - KITDOC V2.0.xlsx")

#DICIONÁRIOS 

rubricas={
    "CCC":5,
    "CCB":6,
    "CHP":8,
    "CBC":9
}

produto={
    "CB":8,
    "CC":9,
    "EMP":10,
    "BD":11,
    "ADTO":13
}

kits={

    #ADIANTAMENTO SALARIAL
    "ADTO":269,
    "ADTO/CPL":270,

    #CARTÃO BENEFICIO - BEM CARTOES
    "CB CBC":228,
    "CB/SQ CBC":228,

    "CB/C CBC":241,
    "CB/CP CBC":241,

    "CB/COMPL CBC":225,
    "CB/CPL CBC":225,

    "CB/C COMPL CBC":243,
    "CB/CP CPL CBC":243,

    #CARTÃO DE CRÉDITO - BEM CARTOES

    "CC CBC":227,
    "CC/SQ CBC":227,

    "CC/C CBC":254,
    "CC/CP CBC":254,

    "CC/COMPL CBC":224,
    "CC/CPL CBC":224,

    "CC/C COMPL CBC":255,
    "CC/CP CPL CBC":255,

    #EMPRESTIMO - BEM CARTOES

    "EMP CBC":226,
    "EMP/CPL CBC":226,
    "EMP/COMPL CBC":226,

    "EMP/C CBC":256,
    "EMP/CP CBC":256,


    #CARTÃO BENEFICIO - CAPITAL
    "CB CCC":213,
    "CB/SQ CCC":213,

    "CB/C CCC":247,
    "CB/CP CCC":247,

    "CB/COMPL CCC":216,
    "CB/CPL CCC":216,

    "CB/C COMPL CCC":246,
    "CB/CP CPL CCC":246,

    #CARTÃO DE CRÉDITO - CAPITAL

    "CC CCC":214,
    "CC/SQ CCC":214,

    "CC/C CCC":249,
    "CC/CP CCC":249,

    "CC/COMPL CCC":217,
    "CC/CPL CCC":217,

    "CC/C COMPL CCC":250,
    "CC/CP CPL CCC":250,

    #EMPRESTIMO - CAPITAL

    "EMP CCC":215,
    "EMP/CPL CCC":215,
    "EMP/COMPL CCC":215,

    "EMP/C CCC":251,
    "EMP/CP CCC":251,

    #CARTÃO BENEFICIO - CLICKBANK
    "CB CCB":212,
    "CB/SQ CCB":212,

    "CB/C CCB":260,
    "CB/CP CCB":260,

    "CB/COMPL CCB":234,
    "CB/CPL CCB":234,

    "CB/C COMPL CCB":263,
    "CB/CP CPL CCB":263,

    #CARTÃO DE CRÉDITO - CLICKBANK

    "CC CCB":219,
    "CC/SQ CCB":219,

    "CC/C CCB":261,
    "CC/CP CCB":261,

    "CC/COMPL CCB":233,
    "CC/CPL CCB":233,

    "CC/C COMPL CCB":264,
    "CC/CP CPL CCB":264,

    #EMPRESTIMO - CLICKBANK

    "EMP CCB":235,
    "EMP/CPL CCB":235,
    "EMP/COMPL CCB":235,

    "EMP/C CCB":262,
    "EMP/CP CCB":262,

    
    #CARTÃO BENEFICIO - HOJE PREVIDENCIA
    "CB CHP":223,
    "CB/SQ CHP":223,

    "CB/C CHP":244,
    "CB/CP CHP":244,

    "CB/COMPL CHP":222,
    "CB/CPL CHP":222,

    "CB/C COMPL CHP":245,
    "CB/CP CPL CHP":245,

    #CARTÃO DE CRÉDITO - HOJE PREVIDENCIA

    "CC CHP":221,
    "CC/SQ CHP":221,

    "CC/C CHP":257,
    "CC/CP CHP":257,

    "CC/COMPL CHP":220,
    "CC/CPL CHP":220,

    "CC/C COMPL CHP":258,
    "CC/CP CPL CHP":258,

    #EMPRESTIMO - HOJE PREVIDENCIA

    "EMP CHP":232,
    "EMP/CPL CHP":232,
    "EMP/COMPL CHP":232,

    "EMP/C CHP":259,
    "EMP/CP CHP":259,
}

#prepara colunas para aceitar texto ou número
preencher["rubrica"]=preencher["rubrica"].astype(object)
preencher["produto"]=preencher["produto"].astype(object)
preencher["kit"]=preencher["kit"].astype(object)

#LOOP PRINCIPAL
for i,linha in preencher.iterrows():

    nome=linha["descricao"]

    if pd.isna(nome):
        continue

    palavras=str(nome).split()

    #RUBRICA
    sigla_rubrica=palavras[-1]

    rubrica_encontrada=rubricas.get(sigla_rubrica)

    if rubrica_encontrada:
        preencher.loc[i,"rubrica"]=rubrica_encontrada
    else:
        preencher.loc[i,"rubrica"]="#ERRO#"


    #PRODUTO
    sigla_produto=str(linha["descricao"]).split()[0].split("/")[0]

    produto_encontrado=produto.get(sigla_produto)

    if produto_encontrado:
        preencher.loc[i,"produto"]=produto_encontrado
    else:
        preencher.loc[i,"produto"]="#ERRO#"

    #IDENTIDADE
    preencher["identidade"]=14

    #SETOR
    preencher["setor"]=2

    #STATUS
    preencher["status"]=1

    #KIT
    sigla_kit=f"{palavras[0]} {palavras[-1]}" 
    
    kit_encontrada=kits.get(sigla_kit) 
    
    if kit_encontrada: preencher.loc[i,"kit"]=kit_encontrada 
    
    else: preencher.loc[i,"kit"]="#ERRO#"

# salvar planilha
preencher.to_excel("MODELO_LOTE - PREENCHIDO.xlsx", index=False)