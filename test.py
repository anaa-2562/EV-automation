import pandas as pd

df = pd.read_csv("HX_Final_20251112_162426.csv")
print([repr(c) for c in df.columns])
