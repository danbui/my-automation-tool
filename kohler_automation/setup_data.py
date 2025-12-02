import pandas as pd

data = {
    'Code': ['K-23475-4-AF']
}

df = pd.DataFrame(data)
df.to_excel('input.xlsx', index=False)
print("Created input.xlsx")
