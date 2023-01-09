from email.header import Header
import pandas as pd
import time

t = time.time()
ml = int(t * 1000)
TimeToCheck = str(ml-2629700000*2)
print(ml)
print(TimeToCheck)



df = pd.read_clipboard(header=None)
print(df)
rows_count = int(len(df.index))


print(rows_count)

rows_count1 = int(len(df.index)-1)
print(rows_count1)




df.insert(0, "SELECT DISTINCT device_id FROM device_info", rows_count*["sm_type='LT' AND sm_value<-30 AND device_timestamp<= now() + INTERVAL 2 MONTH AND device_id="], True)
df.insert(2, " ", rows_count*["OR"], True)

cell_val = df.iloc[rows_count1, 1]
new_cell_val = str(cell_val) + ";"


df.iloc[rows_count1, 2]=" "
df.iloc[rows_count1, 1]=new_cell_val

su = df.rename(columns={df.columns[1]: "WHERE"})

print(su)

su.to_clipboard(index=False)