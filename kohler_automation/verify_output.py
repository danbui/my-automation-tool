import pandas as pd

try:
    df = pd.read_excel('output.xlsx')
    print(df.to_string())
except Exception as e:
    print(f"Error reading output.xlsx: {e}")
