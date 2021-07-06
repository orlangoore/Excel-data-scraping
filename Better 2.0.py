import openpyxl
import string
from os import walk
# C:\Users\itkac\Desktop\lenina_hueta\dauni\Zborovskiy, Yury 11.2019
mypath = r"C:\Users\itkac\Desktop\lenina_hueta\dauni\Zborovskiy, Yury 11.2019"

filenames = next(walk(mypath))

def do_the_work(file_name):
    wb1 = openpyxl.load_workbook(r"C:\Users\itkac\Desktop\lenina_hueta\dauni\Zborovskiy, Yury 11.2019" + "\\" + file_name)
    shoot1 = wb1.get_sheet_names()[0]

    sheet1 = wb1.get_sheet_by_name(shoot1)


    wb2 = openpyxl.load_workbook(r"C:\Users\itkac\Desktop\lenina_hueta\dauni\test.xlsx")
    shoot2 = wb2.get_sheet_names()[0]

    sheet2 = wb2.get_sheet_by_name(shoot2)

    list_of_things = [("C","F"),("I","L"),("Q","S"),("W","X"),("AA","AD"),("AG","AH"),("AM","AN")]




    days = []

    start_pos = 9


    for item in list_of_things:
        if "Kalenderwoche" in sheet1[f'C{start_pos}'].value and sheet1[f'C{start_pos+1}'].value == "Montag":
            for i in range(start_pos, 50):
                if sheet1[f'{item[0]}{i}'].value and sheet1[f'{item[0]}{i}'].value[0] in string.printable:
                    if sheet1[f'{item[0]}{i}'].value == "Zeit" and sheet1[f'{item[0]}{i+1}'].value[0] in string.printable:
                        days.append([sheet1[f'{item[0]}{i-1}'].value,sheet1[f'{item[1]}{i-1}'].value,sheet1[f'{item[0]}{i+1}'].value.split('-')])
                        try:
                            if sheet1[f'{item[0]}{i}'].value == "Zeit" and sheet1[f'{item[0]}{i+2}'].value[0] in string.printable:
                                days.append([sheet1[f'{item[0]}{i - 1}'].value, sheet1[f'{item[1]}{i - 1}'].value, sheet1[f'{item[0]}{i + 2}'].value.split('-')])
                        except:
                            pass
                        try:
                            if sheet1[f'{item[0]}{i}'].value == "Zeit" and sheet1[f'{item[0]}{i+3}'].value[0] in string.printable:
                                days.append([sheet1[f'{item[0]}{i - 1}'].value, sheet1[f'{item[1]}{i - 1}'].value, sheet1[f'{item[0]}{i + 3}'].value.split('-')])
                        except:
                            pass






    for day in days:
        try:
            day[2][0] = day[2][0].split(":")
        except:
            pass
        try:
            day[2][1] = day[2][1].split(":")
        except:
            pass

    final_days = sorted(days,key=lambda x:x[1])
    print(final_days)


    number = 1
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
        number+=1


    wb2.save(r"C:\Users\itkac\Desktop\lenina_hueta\dauni\processed_data\scrape_"+file_name)




for item in filenames[2]:
    do_the_work(item)


