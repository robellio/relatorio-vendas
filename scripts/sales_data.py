from openpyxl import load_workbook
import pandas as pd
import os 
from pathlib import Path 

BASE_DIR = Path(__file__).parent.parent 
DATA_DIR = BASE_DIR /'data/raw_data' 
OUTPUT_DIR = BASE_DIR /'data/processed_data'

sales_files = [
    DATA_DIR / 'Meganium_Sales_Data_-_AliExpress.csv',
    DATA_DIR / 'Meganium_Sales_Data_-_Shopee.csv',
    DATA_DIR / 'Meganium_Sales_Data_-_Etsy.csv'
]

# Cria a pasta de saída caso não exista
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

def red_delimited_file(path):
    return pd.read_csv(
        path,
        sep='[,|\t]',
        engine='python',
        encoding='utf-8'
    )

# Combinar os arquivos 
dfs = [red_delimited_file(file) for file in sales_files]
combined_files = pd.concat(dfs, ignore_index=True)

output_file = OUTPUT_DIR / 'Meganium_Sales_Data.xlsx'

combined_files.to_excel(output_file, index=False, sheet_name='Vendas') 

# Ajustar largura das colunas
def auto_fit_columns(excel_file, sheet_name='Vendas'):
   
    wb = load_workbook(excel_file)
    ws = wb[sheet_name]

    for col in ws.columns:
        max_length = 0
        column = col[0].column_letter

        # Encontrar o conteúdo mais longo na coluna
        for cell in col:
            try:
                if len(str(cell.value)) > max_length:
                    max_length = len(str(cell.value))
            except:
                pass 

        adjusted_width = (max_length + 2)
        ws.column_dimensions[column].width = adjusted_width        

    wb.save(excel_file)

auto_fit_columns(output_file) 

print(f'''Arquivo consolidado salvo em: {output_file}
    Total de registros combinados: {len(combined_files):,}
    ''')