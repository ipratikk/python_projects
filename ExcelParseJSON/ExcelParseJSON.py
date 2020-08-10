# Author - Pratik Goel (https:://linkedin.com/in/ipratikk)

import os
import sys
import xlrd
import json

from datetime import datetime


class ExcelParse():
    def __init__(self,file_path,headers,items):
        self.file_path = file_path
        self.utf_data = None
        self.header_list = headers
        self.item_column_list = items
        self.data_list = {i:'' for i in headers}    #initialising the data_list dictionary
        self.data_list['Items']= []

    def open(self):
        myMap = []      #to store the encoded data from the excel file by row

        if not os.path.isfile(self.file_path):
            raise Exception("File Not Found")

        workbook = xlrd.open_workbook(self.file_path)
        sheet = workbook.sheet_by_index(0)

        for row in range(sheet.nrows):
            myMap.append(sheet.row_values(row))

        self.utf_data = self.cleanData(myMap)


    def cleanData(self,data_map):
        utf_map = []    #to store the utf-8 encoded data

        for row in data_map:
            utf_row = []

            for ele in row:
                utf_str = unicode(ele).encode("UTF-8")

                if utf_str.count('-')>=10:          #EOF condition
                    return utf_map

                utf_row.append(utf_str)

            if utf_row.count("")!=len(utf_row): utf_map.append(utf_row)      #To remove blank rows from data

        return utf_map


    def parse_headers(self):

        for row in self.utf_data:
            for idx in range(len(row)):
                data = row[idx]

                # Check for the condition cell value = 'Name:<string>'
                if data.count('Name')==1:
                    name_data = data.split(':')
                    self.data_list[name_data[0]]=name_data[1]


                # only store data if the header is in the user-defined header list
                if data in self.header_list:

                    if row[idx+1] not in self.header_list:
                        self.data_list[data] = row[idx+1]

                    if data=="Date":
                        date = int(float(row[idx+1]))                   #Converting the excel float date value to integer value
                        self.data_list[data] = self.convert_Date(date)

        #Checking for all expected fields and raising error accordingly
        self.check_expected_fields(self.data_list,self.header_list,"header")


    #Converting date from numemric excel date value to yyyy-mm-dd format
    def convert_Date(self,excel_date):
        python_date = datetime(*xlrd.xldate_as_tuple(excel_date, 0))
        return python_date.strftime('%Y-%m-%d')


    def parse_items(self):
        find_items_start = 0            #to store the row index where we first encounter 'LineNumber'

        for row in self.utf_data:
            for data in row:

                #Find the 1st occurence of Item headers
                if data in self.item_column_list:
                    find_items_start = self.utf_data.index(row)
                    break

        # To store the relative column index of the keys present in the 1st row where we encounter 'LineNumber'
        item_map = {idx:val for idx,val in enumerate(self.utf_data[find_items_start],0)}


        for rows in self.utf_data[find_items_start+1:]:
            tmp_map = {}

            for col in range(len(rows)):
                ele = item_map[col]

                #Check if the column header is in the specified user input column headers and parse data accordingly
                if ele in self.item_column_list:
                    tmp_map[ele]=rows[col]

            self.check_expected_fields(tmp_map,self.item_column_list,"col")

            self.data_list['Items'].append(tmp_map)


    def check_expected_fields(self,data_list,field_list,type):
        for field in field_list:
            try:
                if field not in data_list or data_list[field]=='':
                    if type=="header":
                        print("Warning : Header Required field '" + field +"' not found\n\n")
                    elif type=="col" :
                        print("CriticalError : The column structure is missing a required field - \n\n"+field)
            except:
                pass


    def parse_data(self):
        self.parse_headers()
        self.parse_items()

    def parseToJSON(self):
        json_object = json.dumps(self.data_list, indent=4)
        return json_object


    #Getter methods to access data in the class

    def get_header_data(self):
        return {x:self.data_list[x] for x in self.data_list if x!='Items'}

    def get_item_data(self):
        return self.data_list['Items']

    def get_all_data(self):
        return self.data_list


class Test:
    def main():
        headers = {'Quote Number','Date','Ship To','Ship From','Name'}      #user-defined header list
        items = {"LineNumber",'PartNumber','Description','Price'}           #user-define items list
        in_file_path = sys.argv[1]                                             #file-path input from command-line as argument
        out_file_path = sys.argv[2]
        xls_par = ExcelParse(in_file_path,headers,items)
        xls_par.open()
        xls_par.parse_data()
        json_parsed = xls_par.parseToJSON()
        print(json_parsed)
        with open (out_file_path,"w") as outfile:
            json.dump(json_parsed,outfile)

    if __name__ == '__main__':
        main()