# Case Accenture: Comparativo de Folhas de Pagamento

Este projeto é um **case desenvolvido para a Accenture** na área de Dados. O objetivo principal é **comparar lançamentos de pagamento** extraídos de dois sistemas distintos e gerar uma planilha de saída com totais e diferenças.

---

## Estrutura do Projeto

- **`app.py`**  
  Contém as funções:
  - `transform_s1(...)`: processa o arquivo do Sistema 1 (layout original).  
  - `transform_s2(...)`: processa o arquivo do Sistema 2 (layout novo).  
  - `transform_both_sistemas(...)`: compila os dois resultados, calcula diferenças e retorna um DataFrame unificado.

- **`requirements.txt`**  
  Lista de dependências necessárias para executar o script.

- **Exemplos de entrada**  
  - `Folha Pag_04-2025 (Sistema 1).xlsx`  
  - `Folha Pag_04-2025 (Sistema 2).xlsx`

- **Saída**  
  - `comparativo_sistemas.xlsx` (gerado pelo script com formatação e tabela no Excel)

---

## Pré‑requisitos

- Python 3.8 ou superior  
- Bibliotecas:
  - pandas  
  - openpyxl  
  - XlsxWriter  

Instale via:
```bash
pip install -r requirements.txt
```

---

## Como usar

1. Coloque as planilhas de **entrada** na mesma pasta do script.  
2. Execute:
   ```bash
   python app.py
   ```
3. O arquivo de saída será salvo como `comparativo_sistemas.xlsx`.

---

## Descrição das Funções

- **`transform_s1(path)`**: lê o layout original, extrai ID, Filial, Descrição, Total (centavos) e Data; ordena por total.  
- **`transform_s2(path)`**: lê o layout novo (header na segunda linha), agrupa e soma lançamentos, retorna os mesmos campos em formato unificado.  
- **`transform_both_sistemas(path1, path2)`**: chama as duas funções acima, faz merge inner, calcula diferença e formata valores de volta para Real com duas casas decimais.

---

## Resultado

O Excel final é uma **tabela estruturada** com:

| ID | Filial   | Descrição           | Total Sis 1 | Total Sis 2 | Diferença | Data       |
|----|----------|---------------------|-------------|-------------|-----------|------------|
| 1  | Filial 1 | Salário             | R$ 58.348,97 | R$ 58.348,97 | R$ 0,00   | 30/04/2025 |
| …  | …        | …                   | …           | …           | …         | …          |

Estilizada com separador de milhares, cabeçalho escuro e linhas em cinza-claro.

---

*Case Data — Accenture*