from email.header import Header
from hashlib import new
from operator import index
import pandas as pd
import pyperclip

string1 = "WHERE delivered_ts IS NOT NULL AND shipment_id="
string2 = "OR"

df = pd.read_clipboard(header=None)
print(df)
rows_count = int(len(df.index))


print(rows_count)

rows_count1 = int(len(df.index)-1)
print(rows_count1)

df.insert(0, "SELECT shipment_id FROM acm_shipments", rows_count*["delivered_ts IS NOT NULL AND shipment_id="], True)
df.insert(2, " ", rows_count*["OR"], True)

cell_val = df.iloc[rows_count1, 1]
new_cell_val = str(cell_val) + ";"


df.iloc[rows_count1, 2]=" "
df.iloc[rows_count1, 1]=new_cell_val

su = df.rename(columns={df.columns[1]: "WHERE"})

print(su)

su.to_clipboard(index=False) 