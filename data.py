import datetime
from collections import defaultdict

# ? use default dict

data_lst: list = [
    {
        "SAP ID": "00000000",
        "Name": "ABC",
        "Time": {
            datetime.date(2023, 4, 13): {
                "Login Time": datetime.datetime(2023, 4, 13, 10, 13, 14),
                "Logout Time": datetime.datetime(2023, 4, 13, 19, 10, 5),
                # "Date": datetime.date(2023, 4, 13),
            },
            datetime.date(2023, 4, 14): {
                "Login Time": datetime.datetime(2023, 4, 14, 12, 13, 14),
                "Logout Time": datetime.datetime(2023, 4, 14, 20, 11, 5),
                # "Date": datetime.date(2023, 4, 14),
            },
        },
        "OD (mins.)": 2559,
    },
    {
        "SAP ID": "11111111",
        "Name": "def",
        "Time": {
            datetime.date(2023, 4, 13): {
                "Login Time": datetime.datetime(2023, 4, 12, 11, 13, 14),
                "Logout Time": datetime.datetime(2023, 4, 13, 15, 10, 5),
                # "Date": datetime.date(2023, 4, 13),
            },
            datetime.date(2023, 4, 14): {
                "Login Time": datetime.datetime(2023, 4, 14, 12, 13, 14),
                "Logout Time": datetime.datetime(2023, 4, 14, 20, 11, 5),
                # "Date": datetime.date(2023, 4, 14),
            },
        },
        "OD (mins.)": 4320,
    },
]


data_lst: list = []
data_dict = {}
data_dict["SAP ID"]="00000000"
data_dict["Name"]="ABC"
data_dict["Time"]=defaultdict(lambda: {"Login Time": '--', "Logout Time": '--'})
data_dict["Time"][datetime.date(2023, 4, 13)]["Login Time"]=datetime.datetime(2023, 4, 13, 10, 13, 14)
data_dict["Time"][datetime.date(2023, 4, 13)]["Logout Time"]=datetime.datetime(2023, 4, 13, 19, 10, 5)
data_dict["Time"][datetime.date(2023, 4, 14)]["Login Time"]=datetime.datetime(2023, 4, 14, 12, 13, 14)
data_dict["Time"][datetime.date(2023, 4, 14)]["Logout Time"]=datetime.datetime(2023, 4, 14, 20, 11, 5)
data_dict["OD (mins.)"]=2559
data_lst.append(data_dict)
data_dict = {}
data_dict["SAP ID"]="11111111"
data_dict["Name"]="def"
data_dict["Time"]=defaultdict(lambda: {"Login Time": '--', "Logout Time": '--'})
data_dict["Time"][datetime.date(2023, 4, 13)]["Login Time"]=datetime.datetime(2023, 4, 12, 11, 13, 14)
data_dict["Time"][datetime.date(2023, 4, 13)]["Logout Time"]=datetime.datetime(2023, 4, 13, 15, 10, 5)
data_dict["Time"][datetime.date(2023, 4, 14)]["Login Time"]=datetime.datetime(2023, 4, 14, 12, 13, 14)
data_dict["Time"][datetime.date(2023, 4, 14)]["Logout Time"]=datetime.datetime(2023, 4, 14, 20, 11, 5)
data_dict["OD (mins.)"]=4320
data_lst.append(data_dict)

print(type(data_dict['Time'][datetime.date(2023, 4, 23)]))


data2_lst: list = [
    {
        "SAP ID": "00000000",
        "Name": "ABC",
        "Time": "Login Time \nLogout Time",
        datetime.datetime(2023, 4, 13, 10, 13, 14): f"{datetime.datetime(2023, 4, 13, 10, 13, 14).strftime('%Y-%m-%d')} \n{datetime.datetime(2023, 4, 13, 19, 10, 5).strftime('%Y-%m-%d')}",
        datetime.datetime(2023, 4, 14, 12, 13, 14):f"{datetime.datetime(2023, 4, 13, 10, 13, 14).strftime('%Y-%m-%d')} \n{datetime.datetime(2023, 4, 13, 19, 10, 5).strftime('%Y-%m-%d')}",
        "OD (mins.)": 2559,
    },
    {
        "SAP ID": "11111111",
        "Name": "dfg",
        "Time": "Login Time \nLogout Time",
        datetime.datetime(2023, 4, 13, 10, 13, 14): f"{datetime.datetime(2023, 4, 13, 10, 13, 14).strftime('%Y-%m-%d')} \n{datetime.datetime(2023, 4, 13, 19, 10, 5).strftime('%Y-%m-%d')}",
        datetime.datetime(2023, 4, 14, 12, 13, 14):f"{datetime.datetime(2023, 4, 13, 10, 13, 14).strftime('%Y-%m-%d')} \n{datetime.datetime(2023, 4, 13, 19, 10, 5).strftime('%Y-%m-%d')}",
        "OD (mins.)": 5550,
    },
]

dates_lst: list = [
    datetime.datetime(2023, 4, 1, 10, 0, 0),
    datetime.datetime(2023, 4, 2, 10, 0, 0),
    datetime.datetime(2023, 4, 3, 10, 0, 0),
    datetime.datetime(2023, 4, 4, 10, 0, 0),
    datetime.datetime(2023, 4, 5, 10, 0, 0),
    datetime.datetime(2023, 4, 6, 10, 0, 0),
    datetime.datetime(2023, 4, 7, 10, 0, 0),
    datetime.datetime(2023, 4, 8, 10, 0, 0),
    datetime.datetime(2023, 4, 9, 10, 0, 0),
    datetime.datetime(2023, 4, 10, 10, 0, 0),
    datetime.datetime(2023, 4, 11, 10, 0, 0),
    datetime.datetime(2023, 4, 12, 10, 0, 0),
    datetime.datetime(2023, 4, 13, 10, 0, 0),
    datetime.datetime(2023, 4, 14, 10, 0, 0),
    datetime.datetime(2023, 4, 15, 10, 0, 0),
    datetime.datetime(2023, 4, 16, 10, 0, 0),
    datetime.datetime(2023, 4, 17, 10, 0, 0),
    datetime.datetime(2023, 4, 18, 10, 0, 0),
    datetime.datetime(2023, 4, 19, 10, 0, 0),
    datetime.datetime(2023, 4, 20, 10, 0, 0),
    datetime.datetime(2023, 4, 21, 10, 0, 0),
    datetime.datetime(2023, 4, 22, 10, 0, 0),
    datetime.datetime(2023, 4, 23, 10, 0, 0),
    datetime.datetime(2023, 4, 24, 10, 0, 0),
    datetime.datetime(2023, 4, 25, 10, 0, 0),
]

dates2_lst: list = [
    datetime.date(2023, 4, 1),
    datetime.date(2023, 4, 2),
    datetime.date(2023, 4, 3),
    datetime.date(2023, 4, 4),
    datetime.date(2023, 4, 5),
    datetime.date(2023, 4, 6),
    datetime.date(2023, 4, 7),
    datetime.date(2023, 4, 8),
    datetime.date(2023, 4, 9),
    datetime.date(2023, 4, 10),
    datetime.date(2023, 4, 11),
    datetime.date(2023, 4, 12),
    datetime.date(2023, 4, 13),
    datetime.date(2023, 4, 14),
    datetime.date(2023, 4, 15),
    datetime.date(2023, 4, 16),
    datetime.date(2023, 4, 17),
    datetime.date(2023, 4, 18),
    datetime.date(2023, 4, 19),
    datetime.date(2023, 4, 20),
    datetime.date(2023, 4, 21),
    datetime.date(2023, 4, 22),
    datetime.date(2023, 4, 23),
    datetime.date(2023, 4, 24),
    datetime.date(2023, 4, 25),
]