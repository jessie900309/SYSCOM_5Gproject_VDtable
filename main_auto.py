import datetime
from constants import *
from linkSQL import sqlConnect
from util import *
import pandas as pd

input_data = "example_VD_OpenData.xlsx"


def main_auto():
    try:
        print(welcome)
        start_day = input("請輸入起始日 : ")
        while True:
            try:
                if start_day != datetime.datetime.strptime(
                    start_day, "%Y-%m-%d"
                ).strftime("%Y-%m-%d"):
                    raise ValueError
                else:
                    break
            except ValueError:
                start_day = input("你輸入的日期格式有誤OHO\n請重新輸入日期 : ")
        date_list = []
        for i in range(15):
            nextDate = (
                datetime.datetime.strptime(start_day, "%Y-%m-%d")
                + datetime.timedelta(days=i)
            ).strftime("%Y-%m-%d")
            date_list.append(nextDate)
        conn, cur = sqlConnect()
        ws, wb = openExcelFile(input_data)
        for x in range(3):
            for i in range(15):
                day = date_list[i]
                for j in range(24):
                    print(
                        "\nnow select ... {} {} {} ".format(vd_id[x], day, time_st[j]),
                        end="",
                    )
                    set_vdid = "SET @vdid = '{}';".format(vd_id[x])
                    cur.execute(set_vdid)
                    set_linkid = "SET @linkid = '{}';".format(link_id[x])
                    cur.execute(set_linkid)
                    set_time_st = "SET @TS ='{} {}';".format(day, time_st[j])
                    cur.execute(set_time_st)
                    set_time_ed = "SET @TE ='{} {}';".format(day, time_ed[j])
                    cur.execute(set_time_ed)
                    # main
                    cur.execute(OuO)
                    fetch_data = cur.fetchall()
                    df = pd.DataFrame(fetch_data)
                    # write data
                    if df.empty:
                        print("查無資料", end="")
                        pass
                    else:
                        # sum
                        Total = df[0].sum(0)
                        cellID = COLUMN[x][i] + str(ROW[j])
                        ws[cellID].value = Total
                # write date
                cellID = COLUMN[0][i] + "1"
                ws[cellID].value = day
        wb.save("output.xlsx")
        conn.close()
        print("\n導入完成OuO")
    except KeyboardInterrupt:
        print("Bye Bye :)")
    except Exception as e:
        print("\n---------------- Error ----------------")
        catchError(e)


if __name__ == "__main__":
    main_auto()
