COLUMN = [
    ["B", "E", "H", "K", "N", "Q", "T", "W", "Z", "AC", "AF", "AI", "AL", "AO", "AR"],
    ["C", "F", "I", "L", "O", "R", "U", "X", "AA", "AD", "AG", "AJ", "AM", "AP", "AS"],
    ["D", "G", "J", "M", "P", "S", "V", "Y", "AB", "AE", "AH", "AK", "AN", "AQ", "AT"],
]
ROW = range(3, 27)

time_st = [
    "00:00:00",
    "01:00:00",
    "02:00:00",
    "03:00:00",
    "04:00:00",
    "05:00:00",
    "06:00:00",
    "07:00:00",
    "08:00:00",
    "09:00:00",
    "10:00:00",
    "11:00:00",
    "12:00:00",
    "13:00:00",
    "14:00:00",
    "15:00:00",
    "16:00:00",
    "17:00:00",
    "18:00:00",
    "19:00:00",
    "20:00:00",
    "21:00:00",
    "22:00:00",
    "23:00:00",
]
time_ed = [
    "00:59:59",
    "01:59:59",
    "02:59:59",
    "03:59:59",
    "04:59:59",
    "05:59:59",
    "06:59:59",
    "07:59:59",
    "08:59:59",
    "09:59:59",
    "10:59:59",
    "11:59:59",
    "12:59:59",
    "13:59:59",
    "14:59:59",
    "15:59:59",
    "16:59:59",
    "17:59:59",
    "18:59:59",
    "19:59:59",
    "20:59:59",
    "21:59:59",
    "22:59:59",
    "23:59:59",
]

"""
VD-43-0090-175-01	南三棧(193上方)
VD-42-0090-166-01	崇德段CMS指示牌
VD-42-0090-157-02	仁水隧道
"""

vd_id = ["VD-43-0090-175-01", "VD-42-0090-166-01", "VD-42-0090-157-02"]
link_id = ["3000900117598U", "3000900116579U", "3000900115608U"]

OuO = "SELECT volume FROM vdlive WHERE vdid = @vdid AND linkid = @linkid AND VehicleType IN ('S', 'L') AND datacollecttime BETWEEN @TS and @TE order by datacollecttime, laneid;"

welcome = """
你好，我是可以將VD資料自動導入excel的小工具OuO
資料將以每小時做統計，取連續十五天
確定使用請依提示輸入參數(起始日期)，離開請按Ctrl+C
範例輸入 : 
    ╔═══════════════════════════╗
    ║                           ║
    ║     Date :  2022-07-19    ║
    ║                           ║
    ╚═══════════════════════════╝
"""
