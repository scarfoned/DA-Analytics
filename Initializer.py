import pandas as pd
from math import sin, cos, sqrt, atan2, radians
import csv


def distance_formula(lat1, lat2, lon1, lon2):
    r = int(6371)
    dlon = lon2 - lon1
    dlat = lat2 - lat1
    a = sin(dlat / 2) ** 2 + cos(lat1) * cos(lat2) * sin(dlon / 2) ** 2
    c = 2 * atan2(sqrt(a), sqrt(1 - a))
    distance = r * c
    print(distance)
    return round(distance, 2)


def check_distance(lat, lon):
    distance_lst = []
    lat_length = len(lat)
    lon_length = len(lon)
    front_lat = lat[0]
    front_lon = lon[0]
    if lat_length == 2 and lon_length == 2:
        x = lat[0]
        y = lat[1]
        k = lon[0]
        t = lon[1]
        return distance_lst.append(distance_formula(x, y, k, t))
    if lat_length in (0, 1):
        return False
    for i, j in enumerate(lat[1:]):
        v = lon[1:][i]
        distance_lst.append(distance_formula(front_lat, j, front_lon, v))
    return check_distance(lat[1:], lon[1:])


def list_tgbid(id):
    print(id)


class Import:
    def __init__(self, xldoc, sheet):
        self.xldoc = xldoc
        self.df = pd.read_excel(self.xldoc, sheet_name=sheet)
        self.dictionary = {}
        self.lst_of_lat = []
        self.lst_of_lon = []
        self.tgb_id = []
        self.distances = []

    def arrange_information(self):
        df = self.df
        for line in df.index:
            self.dictionary[df['TGB ID'][line]] = list([df['Latitude'][line], df['Longitude'][line]])
        for i in self.dictionary.values():
            self.lst_of_lat.append(radians(i[0]))
            self.lst_of_lon.append(radians(i[1]))
        for t in self.dictionary.keys():
            self.tgb_id.append(t)

    def distance_of_two_TGBs(self):
        length_of_dictionary = []
        for i in range(len(self.dictionary)):
            length_of_dictionary.append(i)
        check_distance(self.lst_of_lat, self.lst_of_lon)
        list_tgbid(self.tgb_id)

    def send_to_CSV(self):
        with open('mycsv.csv', 'w') as f:
            w = csv.DictWriter(f, self.dictionary.keys())
            w.writeheader()
            w.writerow(self.dictionary)


def open_doc():
    loc1 = str(r'C:\Users\SCA92965\OneDrive - Black & Veatch\Desktop\\')
    file1 = str('TGB.xlsx')
    sheet = str('TGB List')
    new_import = Import((loc1 + file1), sheet)
    new_import.arrange_information()
    new_import.distance_of_two_TGBs()


open_doc()
