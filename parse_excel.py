import pandas as pd
from pathlib import Path  # Example of unused import Codex could clean up

def parse_excel(file_path):
    df = pd.read_excel(file_path)
    # Example of applymap Codex could suggest replacing
    df = df.applymap(str.strip)
    return df

