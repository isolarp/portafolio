import io
import zipfile
import requests
import pandas as pd

URL = "https://www.mercadopublico.cl/Portal/att.ashx?id=5"

resp = requests.get(URL, timeout=30)
resp.raise_for_status()
zip_bytes = io.BytesIO(resp.content)

z = zipfile.ZipFile(zip_bytes)
names = z.namelist()

excel_exts = ('.xls', '.xlsx', '.xlsm', '.xlsb')
dfs = {}

for name in names:
    lname = name.lower()
    if any(lname.endswith(ext) for ext in excel_exts):
        try:
            data = z.read(name)
            bio = io.BytesIO(data)
            if lname.endswith('.xls'):
                engine = 'xlrd'
            elif lname.endswith('.xlsb'):
                engine = 'pyxlsb'
            else:
                engine = 'openpyxl'

            df = pd.read_excel(bio, sheet_name=0, engine=engine, header=None)

            if df.shape[1] == 0 or df.shape[0] <= 7:
                df_clean = pd.DataFrame()
            else:
                df = df.iloc[:, 1:]
                df = df.iloc[7:].reset_index(drop=True)

                if df.shape[0] == 0:
                    df_clean = pd.DataFrame()
                else:
                    new_header = df.iloc[0].astype(str).str.strip()
                    df = df.iloc[1:].reset_index(drop=True)
                    df.columns = new_header
                    df_clean = df.dropna(axis=1, how='all')

            dfs[name] = df_clean
            print(f"Procesado: {name} -> shape={df_clean.shape}")
        except Exception as e:
            print(f"Error procesando {name}: {e}")

if not dfs:
    print("No se encontraron archivos Excel en el zip.")
elif len(dfs) == 1:
    name, df = next(iter(dfs.items()))
    print(f"\nEncontrado un Excel limpio: {name} -> shape={df.shape}")
    print(df.head().to_string(index=False))
else:
    print(f"\nEncontrados {len(dfs)} archivos Excel (limpios):")
    for name, df in dfs.items():
        print("-" * 40)
        print(f"{name} -> shape={df.shape}")
        print(df.head().to_string(index=False))