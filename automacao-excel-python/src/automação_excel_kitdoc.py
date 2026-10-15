import re
import pandas as pd

preencher = pd.read_excel("modelos/MODELO_LOTE - KITDOC V2.0.xlsx")

# - - - - - - - - - - DICIONÁRIOS - - - - - - - - - -
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

    # - - - ADIANTAMENTO SALARIAL - - -
    "ADTO CBC":269,
    "ADIANTAMENTO CBC":269,

    "ADTO CCC":269,
    "ADIANTAMENTO CCC":269,

    "ADTO CCB":269,
    "ADIANTAMENTO CCB":269,

    "ADTO CHP":269,
    "ADIANTAMENTO CHP":269,
    
    "ADTO/CPL CBC":270,
    "ADIANTAMENTO/CPL CBC":270,

    "ADTO/CPL CCC":270,
    "ADIANTAMENTO/CPL CCC":270,

    "ADTO/CPL CCB":270,
    "ADIANTAMENTO/CPL CCB":270,

    "ADTO/CPL CHP":270,
    "ADIANTAMENTO/CPL CHP":270,

    "ADTO/C CBC":276,
    "ADIANTAMENTO/C CBC":276,

    "ADTO/C CCC":276,
    "ADIANTAMENTO/C CCC":276,

    "ADTO/C CCB":276,
    "ADIANTAMENTO/C CCB":276,

    "ADTO/C CHP":276,
    "ADIANTAMENTO/C CHP":276,

    # - - - BENS DURAVEIS - BEM CARTÕES- - -

    "BD CBC":226,

    "BD/CP CBC":256,
    "BD/C CBC":256,


    # - - - CARTÃO BENEFICIO - BEM CARTÕES - - -

    "CB CBC":228,
    "CB/SQ CBC":228,

    "CB/C CBC":241,
    "CB/CP CBC":241,

    "CB/COMPL CBC":225,
    "CB/CPL CBC":225,

    "CB/C COMPL CBC":243,
    "CB/CP CPL CBC":243,
    "CB/C CPL CBC":243,


    # - - - CARTÃO DE CRÉDITO - BEM CARTÕES - - -

    "CC CBC":227,
    "CC/SQ CBC":227,

    "CC/C CBC":254,
    "CC/CP CBC":254,

    "CC/COMPL CBC":224,
    "CC/CPL CBC":224,

    "CC/C COMPL CBC":255,
    "CC/CP CPL CBC":255,
    "CC/C CPL CBC":255,


    # - - - EMPRESTIMO - BEM CARTÕES - - -

    "EMP CBC":226,
    "EMP/CPL CBC":226,
    "EMP/COMPL CBC":226,

    "EMP/C CBC":256,
    "EMP/CP CBC":256,


    # - - - BENS DURAVEIS - CAPITAL - - -

    "BD CCC":215,

    "BD/CP CCC":251,
    "BD/C CCC":251,


    # - - - CARTÃO BENEFICIO - CAPITAL - - -
    "CB CCC":213,
    "CB/SQ CCC":213,

    "CB/C CCC":247,
    "CB/CP CCC":247,

    "CB/COMPL CCC":216,
    "CB/CPL CCC":216,

    "CB/C COMPL CCC":246,
    "CB/CP CPL CCC":246,
    "CB/C CPL CCC":246,


    # - - - CARTÃO DE CRÉDITO - CAPITAL - - -

    "CC CCC":214,
    "CC/SQ CCC":214,

    "CC/C CCC":249,
    "CC/CP CCC":249,

    "CC/COMPL CCC":217,
    "CC/CPL CCC":217,

    "CC/C COMPL CCC":250,
    "CC/CP CPL CCC":250,
    "CC/C CPL CCC":250,

    # - - - EMPRESTIMO - CAPITAL - - -

    "EMP CCC":215,
    "EMP/CPL CCC":215,
    "EMP/COMPL CCC":215,

    "EMP/C CCC":251,
    "EMP/CP CCC":251,


    # - - - BENS DURAVEIS - CLICKBANK - - -

    "BD CCB":235,

    "BD/CP CCB":262,
    "BD/C CCB":262,


    # - - - CARTÃO BENEFICIO - CLICKBANK - - -

    "CB CCB":212,
    "CB/SQ CCB":212,

    "CB/C CCB":260,
    "CB/CP CCB":260,

    "CB/COMPL CCB":234,
    "CB/CPL CCB":234,

    "CB/C COMPL CCB":263,
    "CB/CP CPL CCB":263,
    "CB/C CPL CCB":263,

    # - - - CARTÃO DE CRÉDITO - CLICKBANK - - -

    "CC CCB":219,
    "CC/SQ CCB":219,

    "CC/C CCB":261,
    "CC/CP CCB":261,

    "CC/COMPL CCB":233,
    "CC/CPL CCB":233,

    "CC/C COMPL CCB":264,
    "CC/CP CPL CCB":264,
    "CC/C CPL CCB":264,


    # - - - EMPRESTIMO - CLICKBANK - - -

    "EMP CCB":235,
    "EMP/CPL CCB":235,
    "EMP/COMPL CCB":235,

    "EMP/C CCB":262,
    "EMP/CP CCB":262,


    # - - - BENS DURAVEIS - HOJE PREVIDENCIA - - -

    "BD CHP":232,

    "BD/CP CHP":259,
    "BD/C CHP":259,


    # - - - CARTÃO BENEFICIO - HOJE PREVIDENCIA - - -

    "CB CHP":223,
    "CB/SQ CHP":223,

    "CB/C CHP":244,
    "CB/CP CHP":244,

    "CB/COMPL CHP":222,
    "CB/CPL CHP":222,

    "CB/C COMPL CHP":245,
    "CB/CP CPL CHP":245,
    "CB/C CPL CHP":245,

    # - - - CARTÃO DE CRÉDITO - HOJE PREVIDENCIA - - -

    "CC CHP":221,
    "CC/SQ CHP":221,

    "CC/C CHP":257,
    "CC/CP CHP":257,

    "CC/COMPL CHP":220,
    "CC/CPL CHP":220,

    "CC/C COMPL CHP":258,
    "CC/CP CPL CHP":258,
    "CC/C CPL CHP":258,


    # - - - EMPRESTIMO - HOJE PREVIDENCIA - - -

    "EMP CHP":232,
    "EMP/CPL CHP":232,
    "EMP/COMPL CHP":232,

    "EMP/C CHP":259,
    "EMP/CP CHP":259,
}

# - - - - - - - - - - PREPARAÇÃO - - - - - - - - - -
siglas_modificadoras={"C", "CP", "CPL", "COMPL"}#lista das partes do meio do nome que mudam o kit

def quebrar(texto):
    partes = re.split(r"[\s/_]+", str(texto).strip())

    partes_limpas = []
    for parte in partes:
        if parte:  #ignora pedaços vazios                   
            partes_limpas.append(parte)
    return partes_limpas


def chave_kit(texto):  #só o que precisa para identificar o kit
    t = quebrar(texto)
    produto_ = t[0]
    rubrica_ = t[-1]

    siglas_encontradas=[]
    for modificador in t[1:-1]: #percorre só o meio do nome
        if modificador in siglas_modificadoras:
            siglas_encontradas.append(modificador)
    modificadores=frozenset(siglas_encontradas)
    return (produto_, modificadores, rubrica_)


#converte o dicionário "kits" para o formato preparado
kit_por_linha = {}
for chave, codigo in kits.items(): #chave = texto do kit, codigo=código do kit
    kit_por_linha[chave_kit(chave)] = codigo

# - - - - - - - - - - PREPARANDO PARA PREENCHER - - - - - - - - - -
for coluna in ["rubrica", "produto", "kit"]:
    preencher[coluna] = preencher[coluna].astype(object)

# - - - - - - - - - - PREENCHENDO COLUNAS FIXAS - - - - - - - - - -
preencher["identidade"] = 14
preencher["setor"] = 2
preencher["status"] = 1

# - - - - - - - - - - LOOP PRINCIPAL - - - - - - - - - -
for i, linha in preencher.iterrows():
    nome = linha["descricao"]
    if pd.isna(nome):
        continue

    partes = quebrar(nome) #quebra o nome em partes

    # RUBRICA
    preencher.loc[i, "rubrica"] = rubricas.get(partes[-1], "#ERRO#")

    # PRODUTO
    preencher.loc[i, "produto"] = produto.get(partes[0], "#ERRO#")

    # KIT
    preencher.loc[i, "kit"] = kit_por_linha.get(chave_kit(nome), "#ERRO#")

preencher.to_excel("modelos/MODELO_LOTE - PREENCHIDO.xlsx", index=False)
