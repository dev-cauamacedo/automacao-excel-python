# 📊 Automação Excel com Python

Automação que preenche planilhas de subida em lote a partir da descrição de cada item, eliminando um processo manual, repetitivo e sujeito a erro.

## 📝 Sobre o Projeto

Desenvolvi esta automação no ambiente de trabalho (Grupo AKRK) para resolver um gargalo real: a subida em lote no sistema era feita manualmente, e o tempo dependia da experiência e da familiaridade de cada pessoa com o sistema.

O script lê uma planilha modelo, interpreta a descrição de cada linha (por exemplo, `CC_CPL_PIAUI_LV_F4_CO_PRD_CCC` ou `CC/CPL CCC`) e preenche automaticamente as colunas de **rubrica**, **produto** e **kit**, além dos campos fixos (identidade, setor e status). Linhas que não se encaixam em nenhuma regra recebem `#ERRO#`, para revisão rápida.

## 📈 Resultados

* **80% mais rápido** na subida em lote do sistema
* **Margem de erro reduzida em mais de 93%** já na primeira versão
* **Praticamente zero erros** após a evolução do código: nas execuções seguintes, nenhuma linha foi marcada com `#ERRO#`
* Após o uso em produção, surgiu uma combinação não prevista (compra complementar com `C` e `CPL` juntos). A coluna `#ERRO#` permitiu identificar o caso rapidamente, e a correção foi tratar `CP` como `C` e `COMPL` como `CPL` na chave do kit.
* Processo que dependia da experiência de quem executava agora leva **cerca de 1 minuto** e já sai pronto

## ⚙️ Como Funciona

1. **Leitura:** o Pandas carrega a planilha modelo.
2. **Quebra da descrição:** com expressões regulares (`re`), o nome é separado por espaço, `/` e `_` de uma só vez, aceitando todos os formatos usados no dia a dia.
3. **Chave do kit:** a partir das partes, o script monta uma chave com três informações: produto (primeira parte), modificadores relevantes como `C` e `CPL` (partes do meio) e rubrica (última parte). Siglas que não alteram o kit, como região ou sufixos, são ignoradas.
4. **Consulta:** dicionários de regras de negócio convertem as siglas nos códigos de rubrica, produto e kit.
5. **Tratamento de erros:** combinações desconhecidas recebem `#ERRO#`, o que permite filtrar rapidamente o que precisa de ajuste.
6. **Saída:** a planilha preenchida é salva em um novo arquivo Excel.

### Exemplo

| Descrição | Produto | Modificadores | Rubrica |
|---|---|---|---|
| `CC_CPL_PIAUI_LV_F4_CO_PRD_CCC` | `CC` | `CPL` | `CCC` |
| `CC/CPL CCC` | `CC` | `CPL` | `CCC` |
| `ADTO/C CBC` | `ADTO` | `C` | `CBC` |

Os dois primeiros formatos diferentes resultam na mesma chave e, portanto, no mesmo kit.

## 🔄 Evolução do Projeto

A primeira versão separava o nome apenas por espaço. Descrições com `_` (como `CC_CPL_PIAUI_LV_F4_CO_PRD_CCC`) ficavam como uma única palavra e não encontravam rubrica nem kit. Reescrevi a leitura do nome com expressões regulares e passei a identificar o kit pelos seus componentes, em vez do texto exato. Isso também reduziu a repetição no dicionário de kits e eliminou os erros nesses formatos.

Outras melhorias:

* Valores fixos (identidade, setor, status) atribuídos uma vez, fora do loop
* Uso de `.get()` com valor padrão `#ERRO#`, evitando falhas silenciosas
* Nomes de variáveis e funções mais claros, para facilitar a manutenção

## 🚀 Tecnologias Utilizadas

* **Linguagem:** Python
* **Bibliotecas:** Pandas, openpyxl, re
* **Ferramentas:** Excel, VS Code

## 🔧 Como Executar

1. Clone o repositório.
2. Instale as dependências:

```bash
pip install pandas openpyxl
```

3. Coloque a planilha modelo na pasta `modelos/`.
4. Execute o script:

```bash
python nome_do_script.py
```

5. Abra `modelos/MODELO_LOTE - PREENCHIDO.xlsx` e filtre as colunas por `#ERRO#` para revisar eventuais exceções.

> Se der `PermissionError` ao salvar, feche o arquivo de saída no Excel e rode novamente.

## 📚 Aprendizados

* Automação de processos e manipulação de dados com Pandas
* Uso de expressões regulares para tratar padrões de texto variados
* Modelagem de regras de negócio com dicionários
* Normalização de dados: reduzir entradas diferentes à mesma chave
* Tratamento de exceções para garantir a confiabilidade do resultado
* Depuração a partir de um caso real que falhava em produção
