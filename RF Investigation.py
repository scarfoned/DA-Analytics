import pandas as pd
import numpy as np
import datetime as dt
import timeit
import csv
from statistics import mean


def remove_dictitems(offdev, dictionary):
    for item in offdev:
        if item in dictionary:
            del dictionary[item]
    return dictionary


class RFcoverage:
    def __init__(self):
        self.start = '10/1/2018'
        self.end = '10/31/2018'
        self.success = {}
        self.average_success = {}
        self.isoffline = {}
        self.offline_devices = []
        self.bad_devices = []


    def import_csv(self):
        try:
            dataframe = pd.read_csv("C:/Users/C125238/Desktop/Doc Central/All RF Coverage.csv", parse_dates=True, error_bad_lines=False, engine='python')
            df2 = pd.read_csv("C:/Users/C125238/Desktop/Doc Central/DeviceList.csv", parse_dates=True, error_bad_lines=False, engine='python')

            dataframe.Date = pd.to_datetime(dataframe.Date)
            start = pd.to_datetime(self.start)
            end = pd.to_datetime(self.end)
            current_month = dt.datetime.today().month
            #df = dataframe.loc[dataframe.Date.dt.month == current_month, : ]
            mask = (dataframe['Date'] >= start) & (dataframe['Date'] <= end)
            df = dataframe.loc[mask]

            self.success = dict(df.groupby("DeviceNumber")["Success"].apply(list).to_dict())
            self.isoffline = dict(zip(df2["Device ID"], df2["In Service"]))

            """Saves dictionary with Device Number as the key and Success as the value in the form of a list ex: {24753: [4, 9, 10, 10]}"""
            return self.success, self.isoffline
        except TypeError as e:
            print("Fix: ", e)


    def filter_dn(self):
        for i in self.isoffline:
            if self.isoffline[i] == False:
                self.offline_devices.append(i)
        remove_dictitems(self.offline_devices, self.success)
        return self.success


    def organize_dns(self):
        """Sends the self.success list to remove_items, which removes the smallest values from the list created in import_csv and leaves two values"""
        for i in self.success:
            if len(self.success[i]) > 2:
                while len(self.success[i]) > 2:
                    self.success[i].remove(min(self.success[i]))
        return self.success


    def average_rtm(self):
        num_greater = 0

        """Gets the mean of the two values in self.success and creates a new dictionary with the device number as key and the average as the value"""
        for i in self.success:
            print(self.success[i])
            self.average_success[i] = mean(self.success[i])

        """Gets the length of the self.average_success dictionary"""
        list_length = len(self.average_success.keys())

        """Checks if the average of each value is greater than 9.5 and if so deliminates that device as having good rf coverage"""
        for j in self.average_success:
            if self.average_success[j] > 8.5:
                num_greater += 1
            else:
                self.bad_devices.append(j)

        self.bad_devices.insert(0, 'RTM #')

        """Gets the average of the devices and the length of the device list"""
        try:
            average = num_greater / list_length
            print(average, "%.0f%%" % (100 * average))
            print("Devices that have bad comms:", self.bad_devices)
            self.bad_devices.insert(1, average)
        except ZeroDivisionError:
            print("Date's are out of range of dataset")

        return self.average_success, self.bad_devices


    def export_excel(self):
        with open('Bad Devices-'+self.start[0]+'.csv', 'w', newline='') as bdexport:
            wr = csv.writer(bdexport)
            wr.writerows(enumerate(self.bad_devices))


def main():
    get = RFcoverage()
    get.import_csv()
    get.filter_dn()
    get.organize_dns()
    get.average_rtm()
    get.export_excel()


if __name__ == "__main__":
    main()