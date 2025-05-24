# Importar bibliotecas
import pandas as pd
from datetime import datetime as dt
from xlsxwriter import Workbook

# Função para Sistema 1

def transform_s1(input_path: str, sheet_name=0) -> pd.DataFrame:
    """
    Lê o Excel do Sistema 1 (layout original) e retorna DataFrame com colunas:
    ID, Filial, Descrição, Total e Data.
    A data é fixa: 2 linhas abaixo na coluna D.
    """
    df = pd.read_excel(input_path, header=None, sheet_name=sheet_name)
    records = []

    for idx, row in df.iterrows():
        cell = row[1]
        if isinstance(cell, str) and cell.strip().lower().startswith('filial'):
            branch_name = cell.strip()
            id_val = row[0]
            id_str = str(int(id_val)) if pd.notna(id_val) else ''

            # Data 2 linhas abaixo (idx+2), coluna D (índice 3)
            date_str = ''
            if idx + 2 < len(df):
                date_val = df.iat[idx+2, 3]
                if pd.notna(date_val):
                    if isinstance(date_val, (pd.Timestamp, dt)):
                        date_str = date_val.strftime('%d/%m/%Y')
                    else:
                        date_str = str(date_val)

            # Cabeçalho "Descrição"
            header_idxs = df[(df.index > idx) & (df[1] == 'Descrição')].index
            if len(header_idxs) == 0:
                continue
            start = header_idxs[0] + 2

            # Coleta linhas até próximo bloco de filial
            for j in range(start, len(df)):
                r = df.iloc[j]
                nxt = r[1]
                if isinstance(nxt, str) and nxt.strip().lower().startswith('filial'):
                    break
                if pd.isna(r[0]):
                    continue
                total_val = r[3]
                # converte para centavos inteiros
                cents = int(round(total_val * 100)) if pd.notna(total_val) else 0
                records.append({
                    'ID': id_str,
                    'Filial': branch_name,
                    'Descrição': r[1],
                    'Total': cents,
                    'Data': date_str
                })

    result_df = pd.DataFrame(records, columns=['ID', 'Filial', 'Descrição', 'Total', 'Data'])
    result_df = result_df.sort_values(by='Total', ascending=False).reset_index(drop=True)
    return result_df


# Função para Sistema 2

def transform_s2(input_path: str, sheet_name=0) -> pd.DataFrame:
    """
    Lê o Excel do Sistema 2 (layout novo) e retorna DataFrame com colunas:
    ID, Filial, Descrição, Total (em centavos) e Data.
    Usa a segunda linha (header=1) para extrair nomes das colunas.
    Agrupa por filial, descrição e data, soma valores.
    """
    df = pd.read_excel(input_path, sheet_name=sheet_name, header=1)

    # Detecta colunas-chave
    id_col = next(c for c in df.columns if 'filial' in c.lower())
    desc_col = next(c for c in df.columns if 'hist' in c.lower())
    date_col = next(c for c in df.columns if 'data' in c.lower())
    valor_col = next(c for c in df.columns if 'valor' in c.lower())

    grouped = df.groupby([id_col, desc_col, date_col], as_index=False)[valor_col].sum()
    records = []
    for _, row in grouped.iterrows():
        id_val = row[id_col]
        id_str = str(int(id_val)) if pd.notna(id_val) else ''
        filial_name = f"Filial {id_str}"

        date_val = row[date_col]
        if pd.notna(date_val):
            if isinstance(date_val, (pd.Timestamp, dt)):
                date_str = date_val.strftime('%d/%m/%Y')
            else:
                date_str = str(date_val)
        else:
            date_str = ''

        total_val = row[valor_col]
        # converte para centavos inteiros
        cents = int(round(total_val * 100)) if pd.notna(total_val) else 0
        records.append({
            'ID': id_str,
            'Filial': filial_name,
            'Descrição': row[desc_col],
            'Total': cents,
            'Data': date_str
        })

    result_df = pd.DataFrame(records, columns=['ID', 'Filial', 'Descrição', 'Total', 'Data'])
    result_df = result_df.sort_values(by='Total', ascending=False).reset_index(drop=True)
    return result_df


# Função compiladora

def transform_both_sistemas(path_s1: str, path_s2: str) -> pd.DataFrame:
    """
    Executa transform_s1 e transform_s2, retornando um DataFrame com colunas:
    ID, Filial, Descrição, Total Sis 1, Total Sis 2, Diferença e Data.
    Ordena primeiro por ID (crescente) e depois por Total Sis 1 (decrescente).
    """
    # Gera cada DataFrame
    df1 = transform_s1(path_s1)
    df2 = transform_s2(path_s2)

    # Padroniza o nome da Filial no df1
    df1['Filial'] = df1['ID'].apply(lambda x: f"Filial {x}")

    # Faz merge apenas nos registros iguais em ambos
    merged = pd.merge(
        df1, df2,
        on=['ID', 'Filial', 'Descrição', 'Data'],
        how='inner',
        suffixes=(' Sis 1', ' Sis 2')
    )

    # Preenche possíveis NaNs e garante inteiros
    merged[['Total Sis 1', 'Total Sis 2']] = merged[['Total Sis 1', 'Total Sis 2']].fillna(0).astype(int)
    # Calcula diferença em centavos
    merged['Diferença'] = merged['Total Sis 1'] - merged['Total Sis 2']

    # Converte centavos para float com 2 casas
    for col in ['Total Sis 1', 'Total Sis 2', 'Diferença']:
        merged[col] = (merged[col] / 100).round(2)

    # Ordena por ID crescente e Total Sis 1 decrescente
    merged = merged.sort_values(by=['ID', 'Total Sis 1'], ascending=[True, False]).reset_index(drop=True)

    return merged[['ID', 'Filial', 'Descrição', 'Total Sis 1', 'Total Sis 2', 'Diferença', 'Data']]

# Gerando df final
df_final = transform_both_sistemas("Folha Pag_04-2025 (Sistema 1).xlsx",
                                   "Folha Pag_04-2025 (Sistema 2).xlsx")

# Caminho e nome do arquivo de saída
output_path = "comparativo_sistemas.xlsx"

# Exporta para Excel com formatação
with pd.ExcelWriter(output_path, engine='xlsxwriter') as writer:
    df_final.to_excel(writer, sheet_name='Comparativo', index=False)
    workbook  = writer.book
    worksheet = writer.sheets['Comparativo']
    
    # Formatos
    header_fmt   = workbook.add_format({
        'bold': True,
        'align': 'left',
        'valign': 'vcenter'
    })
    text_fmt     = workbook.add_format({
        'align': 'left',
        'valign': 'vcenter'
    })
    currency_fmt = workbook.add_format({
        'num_format': 'R$ #,##0.00',
        'align': 'left',
        'valign': 'vcenter'
    })
    
    # Reescreve o cabeçalho com o novo formato
    for col_num, header in enumerate(df_final.columns):
        worksheet.write(0, col_num, header, header_fmt)
    
    # Ajusta colunas
    worksheet.set_column('A:A', 8,  text_fmt)       # ID
    worksheet.set_column('B:B', 15, text_fmt)       # Filial
    worksheet.set_column('C:C', 30, text_fmt)       # Descrição
    worksheet.set_column('D:F', 15, currency_fmt)   # Totais + Diferença
    worksheet.set_column('G:G', 12, text_fmt)       # Data
    
    # Converte todo o range em uma Tabela do Excel com estilo claro
    max_row, max_col = df_final.shape
    worksheet.add_table(0, 0, max_row, max_col-1, {
        'columns': [{'header': h} for h in df_final.columns],
        'style': 'Table Style Medium 1',   # linhas alternadas: branco e cinza claro
        'header_row': True,
        'autofilter': True
    })

print(f"Arquivo Excel formatado salvo em: {output_path}")