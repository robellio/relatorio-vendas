import pandas as pd
import os 
from pathlib import Path 

BASE_DIR = Path(__file__).parent.parent 
DATA_DIR = BASE_DIR /'data/raw_data' 
OUTPUT_DIR = BASE_DIR /'data/processed_data'

siles_files = [
    DATA_DIR / 'Meganium_Sales_Data_-_AliExpress.csv',
    DATA_DIR / 'Meganium_Sales_Data_-_Shopee.csv',
    DATA_DIR / 'Meganium_Sales_Data_-_Etsy.csv'
]

# Cria a pasta de saída caso não exista
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

# Combinar os arquivos 
combined_files = pd.concat([pd.read_csv(file) for file in siles_files], ignore_index=True)

output_file = OUTPUT_DIR / 'Meganium_Sales_Data.csv'

combined_files.to_csv(output_file, index=False) 

print(f'''Arquivo consolidado salvo em: {output_file}
    Total de registros combinados: {len(combined_files):,}
    ''')