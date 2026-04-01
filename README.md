# Contagem de Cartas — Autos de Infração (SERASA / CCOBI)

Aplicação para calcular **quantas cartas** seriam necessárias a partir de uma **planilha Excel** de autos de infração, seguindo a regra de negócio acordada no projeto **CCOBI – SERASA**.

---

## Regra de negócio (resumo)

- A contagem é feita **por dia** (não se acumula entre dias diferentes).  
- Para cada **combinação (data de inscrição, CPF/CNPJ)**:
  - Contam-se os **autos únicos** (pelo número do auto; duplicatas na planilha são removidas, mantendo a primeira ocorrência).
  - **1 a 10 autos** no mesmo dia, para o mesmo CPF/CNPJ → **1 carta**  
  - **11 a 20** → **2 cartas**  
  - **21 a 30** → **3 cartas**  
  - E assim por diante: **cartas = teto(autos ÷ 10)** (equivalente a `ceil(autos / 10)`).

Ou seja: a cada bloco de até 10 autos no mesmo dia e mesmo documento, soma-se uma carta.

---

## Funcionalidades

| Recurso | Descrição |
|---------|-----------|
| **Interface gráfica** (`contagem_cartas_gui.py`) | Seleção de planilha, pasta de saída, filtro por **todas as datas**, **mês/ano** ou **datas específicas**, processamento em thread e abertura da pasta de resultado. |
| **Linha de comando** (`contagem_cartas.py`) | Processamento direto sem GUI. |
| **Configuração de colunas** | Arquivo `config_colunas.json` com os nomes exatos das colunas no Excel. |
| **Saída** | Planilha `*_resultado_cartas.xlsx` com abas **Resumo**, **Por dia**, **Por CNPJ**, **Por CNPJ por dia**. |

---

## Requisitos

- **Python** 3.9+ (recomendado)  
- Dependências:

```bash
pip install -r requirements.txt
```

- `pandas` — leitura e agregação  
- `openpyxl` — leitura/escrita `.xlsx`

---

## Formato da planilha de entrada

O arquivo deve ser **`.xlsx`**. As colunas (nomes podem ser ajustados no JSON) **padrão** são:

| Conceito | Nome padrão da coluna |
|----------|------------------------|
| Identificador do auto | `Número do auto` |
| Data | `Data de inscrição` |
| Documento | `CPF/CNPJ` |

### `config_colunas.json`

```json
{
  "numero_auto": "Número do auto",
  "data_inscricao": "Data de inscrição",
  "cpf_cnpj": "CPF/CNPJ"
}
```

Altere os valores para bater **exatamente** com o cabeçalho da sua planilha (após `strip` dos nomes).

### Formatos aceitos para data

- Texto `dd/mm/aaaa` ou `dd-mm-aaaa`  
- Objeto `datetime`  
- **Número serial do Excel** (dias desde 30/12/1899), em faixa plausível  

Linhas **sem data válida** ou **sem CPF/CNPJ** normalizado são ignoradas no cálculo; o resumo indica quantas linhas foram ignoradas e quantos autos duplicados foram removidos.

### CPF/CNPJ

Comparação usa **apenas dígitos** (pontuação é removida para agrupamento).

---

## Como executar

### Modo gráfico (recomendado)

```bash
python contagem_cartas_gui.py
```

Ou use `executar.bat` se existir na pasta.

1. **Selecionar** a planilha `.xlsx`.  
2. Opcional: escolher **pasta de saída** (senão o arquivo vai para a mesma pasta da entrada).  
3. Opcional: **filtrar** por mês ou datas após carregar a planilha.  
4. Clicar em processar e aguardar o resumo (total de cartas, total de autos únicos, etc.).

### Modo linha de comando

```bash
python contagem_cartas.py "C:\caminho\planilha.xlsx"
python contagem_cartas.py "C:\caminho\planilha.xlsx" "C:\pasta_saida"
```

Gera `planilha_resultado_cartas.xlsx` na pasta de saída ou ao lado da entrada.

---

## Saída (abas do Excel)

1. **Resumo** — totais de cartas, autos únicos, linhas ignoradas, datas inválidas, duplicados removidos.  
2. **Por dia** — soma de cartas e autos por data (`dd/mm/aaaa`).  
3. **Por CNPJ** — totais no período (ou no filtro) por documento.  
4. **Por CNPJ por dia** — detalhamento dia a dia por documento.

---

## Estrutura do projeto

| Arquivo | Papel |
|---------|--------|
| `contagem_cartas.py` | Lógica: leitura, normalização, agrupamento, exportação |
| `contagem_cartas_gui.py` | Interface Tkinter |
| `config_colunas.json` | Mapeamento de colunas |
| `requirements.txt` | Dependências |
| `executar.bat` | Atalho Windows para a GUI |

---

## Validação interna

O código compara a soma das cartas na aba **Por dia** com o total geral e imprime aviso se houver divergência.

---

## Privacidade

Planilhas de autos podem conter **dados pessoais**. Não versionar arquivos reais no Git; use apenas amostras anonimizadas.

---

## Contexto

Projeto **CCOBI – SERASA** — automação e análise de autos de infração ANTT.
