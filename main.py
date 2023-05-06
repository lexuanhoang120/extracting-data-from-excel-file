import openpyxl
import os, glob
import pandas as pd
import unidecode
import numpy as np
#path of folder contain the data
path = 'Lan2/Thang 06'

#function to remove accent
def remove_accent(text):
    return unidecode.unidecode(text)

#function to add the information to the dataframe
def add_info(df, store, date, number_tax, company, address, time ,price, name, phone_number,note,more):
    #Save the information to append to DataFrame
    information_new = {'Store':store,'Date':date,'Number_tax':number_tax,'Company':company,'Address':address,'Time':time,'Total_Price':price,"Name":name,'Number_Phone':phone_number,'Note':note,"More":more}
    df_new = df.append(information_new,ignore_index=True)
    return df_new
#columns of the new data
columns = ['Store','Date','Number_tax','Company','Address',"Time","Total_Price","Name","Number_Phone","Note","More"]

df_new = pd.DataFrame([],columns=columns)
#list of stores
stores = ['BINH PHU',
 'CACH MANG THANG 8',
 'CAU CHU Y',
 'DONG DEN',
 'DUONG BA TRAC',
 'MAC DINH CHI',
 'NGUYEN SON',
 'NGUYEN TRAI',
 'NGUYEN TRI PHUONG',
 'NHI THIEN DUONG',
 'PHAN CHU TRINH',
 'PHAN XICH LONG',
 'THANH THAI',
 'TO KY',
 'TRAN BINH TRONG',
 'TRAN NAO',
 'UNG VAN KHIEM',
 'VUNG TAU',
 'XO VIET NGHE TINH']

districs = [
    "QUAN 1",
    "QUAN 2",
    "QUAN 3",
    "QUAN 4",
    "QUAN 5",
    "QUAN 6",
    "QUAN 7",
    "QUAN 8",
    "QUAN 9",
    "QUAN 10",
    "QUAN 11",
    "QUAN 12",
    "THU DUC",
    "BINH TAN",
    "BINH THANH",
    "GO VAP",
    "PHU NHUAN",
    "TAN BINH",
    "TAN PHU",
    "BINH CHANH",
    "CAN GIO",
    "CU CHI",
    "HOC MON",
    "NHA BE"
]

error=[]

for file in glob.glob(os.path.join(path, '*.xlsx')):
    try:
        xl = pd.ExcelFile(file)
        sheets = len(xl.sheet_names)
        #define date
        date = str(file.split("/")[-1])[13:-5]
        for sheet in range(0,sheets):
            df = pd.read_excel(file,sheet_name=sheet)
            
            if df.empty:
                # print(f"Sheet {sheet + 1} is empty.")
                break
            # print(f"Sheet {sheet + 1} have information")
            df = df.values
            
            
            df = pd.DataFrame(df[1:],columns=df[0])
            df = df.loc[:, ~df.columns.duplicated()]
            lst = []
            for i in df.columns:
                lst.append(remove_accent(str(i)))
            df.columns = lst
            for col in df.columns:
                if col =="nan":
                    del df["nan"]
                # if col =="Ghi chu":
                #     del df["Ghi chu"]

            # processing the data
            for index in df.index:
                flag = False
                # if df[df.index==index]['So hoa don'] == None:
                #     continue
                shd = df[df.index==index]['So hoa don']
                name_store = remove_accent(str(shd)).upper()
                # define store
                for sto in stores:
                    if (sto in name_store):
                        flag = True
                        store = sto
                        break
                if flag:
                    continue
                
                number_tax =  list(df[df.index ==index]["Ma so thue"].values)[0]
                # if "'" in number_tax:
                #     number_tax = number_tax[1:-1]
                company = list(df[df.index ==index]["Doanh nghiep"].values)[0]
                address = list(df[df.index ==index]["Dia chi cong ty"].values)[0]
                # if df[df.index == index +1]["Dia chi cong ty"] == None:
                #     price = list((df[df.index ==index + 2]["So tien"].values*(1+8/100)))[0]
                # else:
                price = list((df[df.index ==index]["So tien"].values*(1+8/100)))[0]
                # price = list(df[df.index ==index]["So tien"].values)[0]
                # price = price[1:-1]
                # if "'" in price:
                #     price = price[1:-1]
                name = list(df[df.index ==index]["Nguoi lien he"].values)[0]
                phone_number = list(df[df.index ==index]["Dien thoai"].values)[0]
                # phone_number = phone_number[1:-1]
                # if "'" in phone_number:
                #     phone_number= phone_number[1:-1]
                note = list(df[df.index ==index]["Dia chi gui"].values)[0]
                if "Ghi chu" in df.columns:
                    more = list(df[df.index ==index]["Ghi chu"].values)[0]
                else:
                    more = None
                if "Gio" in df.columns:
                    time = list(df[df.index ==index]["Gio"].values)[0]
                else:
                    time = None
                
                df_new = add_info(df_new, store, date, number_tax, company, address, time ,price, name, phone_number,note,more)
    except:
        error.append(file)  
for err in error:
    print(err)

df = df_new
#filter the data
for index in df.index:
    a = list(df[df.index == index]["Address"].values)[0]
    if a == None:
        df = df.drop([index])
for index in df.index:
    a = list(df[df.index == index]["Total_Price"].values)[0]
    if (a == 0) or (a == None):
        df = df.drop([index])

for i in df.index:
    if i == 0:
        continue
    if df[df.index==i].Date.values == df[df.index==(i-1)].Date.values:
        if df[df.index==i].Number_tax.values == df[df.index==i-1].Number_tax.values:
            df.Total_Price.loc[df.index ==i] = int(float(df[df.index==i].Total_Price) +float(df[df.index==i-1].Total_Price))
            df = df.drop([i-1])

df.to_csv("data_month6.csv")