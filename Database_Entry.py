import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from datetime import datetime, date, timedelta


class Import:
    def __init__(self):
        """Import csv file and other variables"""
        self.content = pd.read_csv('Devices.csv', header=0, index_col=0, encoding='utf-8', parse_dates=True, error_bad_lines=False)
        self.eight_weeks = pd.datetime.today() - timedelta(56)
        self.this_year = pd.datetime.today().year

        """Entire Dataset"""
        self.df = pd.DataFrame(self.content)
        self.df['Online Date'] = pd.to_datetime(self.df['Online Date'])
        self.df['Week Number'] = self.df['Online Date'].dt.week

        """Only devices that have been turned online"""
        self.online_date = {}
        """own stands for Online Week Numbers and cown stands for Count of Online Week Numbers"""
        self.week_num = []
        self.cown = []

        """Devices used to calculate year to date"""
        self.year_to_date = []


    def format_online_devices(self):
        onlinedf = self.df
        onlinedf = onlinedf.dropna(axis=0, subset=['Online Date'])
        onlinedf = onlinedf[(onlinedf['Online Date'].dt.year >= self.this_year)]
        onlinedf = onlinedf[(onlinedf['Online Date'] >= self.eight_weeks)]
        self.online_date = dict(onlinedf.groupby('Week Number')['Week Number'].count())
        self.week_num = list(self.online_date.keys())
        self.cown = list(self.online_date.values())
        return self.online_date, onlinedf, self.week_num, self.cown


    def format_installed_ytd(self):
        ytd = []
        df = self.df
        df = df[(df['Online Date'].dt.year >= self.this_year)]
        self.year_to_date = list(np.cumsum(self.cown))
        return self.year_to_date


    def plot_onlinedf(self):
        xvalue = self.week_num
        yvalue = self.cown
        zvalue = self.year_to_date

        plt.style.use('seaborn-dark')
        plt.figure(1, figsize=(12, 9))

        ax = plt.subplot(131)
        ax.bar(xvalue, yvalue, color='b')
        plt.title('Devices Installed & Online')
        plt.xlabel('Week Number')
        plt.ylabel('Number of Devices')
        rects1 = ax.patches
        labels1 = [yvalue[i] for i in range(len(yvalue))]
        for rect, label in zip(rects1, labels1):
            height = rect.get_height()
            ax.text(rect.get_x() + rect.get_width() / 2, height + .5, label, ha='center', va='bottom')

        ap = plt.subplot(132)
        ap.bar(xvalue, zvalue, color='g')
        plt.title('Devices Installed YTD')
        plt.xlabel('Week Number')
        plt.ylabel('Number of Devices')
        rects2 = ap.patches
        labels2 = [zvalue[i] for i in range(len(yvalue))]
        for rect, label in zip(rects2, labels2):
            height = rect.get_height()
            ap.text(rect.get_x() + rect.get_width() / 2, height + .5, label, ha='center', va='bottom')

        plt.show()


def run_program():
    of = Import()
    of.format_online_devices()
    of.format_installed_ytd()
    of.plot_onlinedf()


if __name__ == '__main__':
    run_program()
