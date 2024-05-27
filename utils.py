import pandas as pd

def readstate_inclusions_from_excel(file_path):
    df = pd.read_excel(file_path)
    return df['Insc. Estadual'].tolist()  # Certifique-se de que o nome da coluna está correto
