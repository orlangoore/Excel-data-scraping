import openpyxl
import string
from os import walk
import concurrent.futures




mypath = r"C:\excel_project\raw_files"
output_path = r"C:\excel_project\processed_files"


class ExcelScraper():
    def __init__(self,source_path, output_path):
        self.source_path = source_path
        self.output_path = output_path
        self.days = []
        self.list_of_columns = [("C","F"),("I","L"),("Q","S"),("W","X"),("AA","AD"),("AG","AH"),("AM","AN")]
        self.start_pos = 9
        self.list_of_filenames = self.get_filenames()

    def __str__(self):
        return f"Reworks timeplans. Source is {self.source_path}, Output is {self.output_path}"

    def __repr__(self):
        return f"From {self.source_path} into {self.output_path}"


    def get_filenames(self):
        list_of_names = []
        for item in next(walk(self.source_path))[2:]:
            for name in item:
                list_of_names.append(name)
        return list_of_names

    def process_files(self,file_name):
        work_book1 = openpyxl.load_workbook(f"{self.source_path}" + "\\" + file_name)
        sheet1 = work_book1.get_sheet_by_name(work_book1.get_sheet_names()[0])

        work_book2 = openpyxl.load_workbook(r"C:\excel_project\test.xlsx")
        sheet2 = work_book2.get_sheet_by_name(work_book2.get_sheet_names()[0])

        for item in self.list_of_columns:
            if "Kalenderwoche" in sheet1[f'C{self.start_pos}'].value and sheet1[f'C{self.start_pos+1}'].value == "Montag":
                for i in range(self.start_pos, 50):
                    if sheet1[f'{item[0]}{i}'].value and sheet1[f'{item[0]}{i}'].value[0] in string.printable:
                        if sheet1[f'{item[0]}{i}'].value == "Zeit" and sheet1[f'{item[0]}{i+1}'].value[0] in string.printable:
                            self.days.append([sheet1[f'{item[0]}{i-1}'].value, sheet1[f'{item[1]}{i-1}'].value, sheet1[f'{item[0]}{i+1}'].value.split('-')])
                            try:
                                if sheet1[f'{item[0]}{i}'].value == "Zeit" and sheet1[f'{item[0]}{i+2}'].value[0] in string.printable:
                                    self.days.append([sheet1[f'{item[0]}{i - 1}'].value, sheet1[f'{item[1]}{i - 1}'].value, sheet1[f'{item[0]}{i + 2}'].value.split('-')])
                            except:
                                pass
                            try:
                                if sheet1[f'{item[0]}{i}'].value == "Zeit" and sheet1[f'{item[0]}{i+3}'].value[0] in string.printable:
                                    self.days.append([sheet1[f'{item[0]}{i - 1}'].value, sheet1[f'{item[1]}{i - 1}'].value, sheet1[f'{item[0]}{i + 3}'].value.split('-')])
                            except:
                                pass

        for day in self.days:
            try:
                day[2][0] = day[2][0].split(":")
            except:
                pass
            try:
                day[2][1] = day[2][1].split(":")
            except:
                pass

        final_days = sorted(self.days, key=lambda x: x[1])

        counter = 1
        for day in final_days:
            try:sheet2[f"A{number}"].value = day[1]
            except:pass
            try:sheet2[f"B{number}"].value = day[0]
            except:pass
            try:sheet2[f"D{number}"].value = day[2][0][0]
            except:pass
            try:sheet2[f"E{number}"].value = day[2][0][1]
            except:pass
            try:sheet2[f"F{number}"].value = day[2][1][0]
            except:pass
            try:sheet2[f"G{number}"].value = day[2][1][1]
            except:pass
            counter += 1

        work_book2.save(f"{self.output_path}\\scrape_" + file_name)

    def do_your_thing(self):
        with concurrent.futures.ThreadPoolExecutor() as exector:
            exector.map(self.process_files, self.list_of_filenames)

scraper = ExcelScraper(mypath, output_path)
scraper.do_your_thing()

