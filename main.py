from constants import *
from linkSQL import sqlConnect
from util import openExcelFile
import pandas as pd

# todo 改變日期(最多連續15天)
date_mon = 9
date_day = range(1, 16)

conn, cur = sqlConnect()
ws, wb = openExcelFile('VD資料統計_20220901_20220915.xlsx')

for x in range(3):
    for i in range(15):
        day = date_day[i]
        for j in range(24):
            print("\nnow select ... {} 2022-{:02d}-{:02d} {}' ".format(vd_id[x], date_mon, day, time_st[j]), end='')
            set_vdid = "SET @vdid = '{}';".format(vd_id[x])
            cur.execute(set_vdid)
            set_linkid = "SET @linkid = '{}';".format(link_id[x])
            cur.execute(set_linkid)
            setTime_st = "SET @TS ='2022-{:02d}-{:02d} {}' ;".format(date_mon, day, time_st[j])
            cur.execute(setTime_st)
            setTime_ed = "SET @TE ='2022-{:02d}-{:02d} {}' ;".format(date_mon, day, time_ed[j])
            cur.execute(setTime_ed)
            # main
            cur.execute(OuO)
            fetch_data = cur.fetchall()
            df = pd.DataFrame(fetch_data)
            # write data
            if df.empty:
                print("查無資料", end='')
                pass
            else:
                # sum
                Total = df[0].sum(0)
                cellID = COLUMN[x][i] + str(ROW[j])
                ws[cellID].value = Total
        # wb.save("output_now.xlsx")
        # ws, wb = openExcelFile('output_now.xlsx')

wb.save("output.xlsx")
conn.close()
