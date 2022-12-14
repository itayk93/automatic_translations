"""""""""
The code was written by Itay Karkaosn
Using mtranslate from - https://github.com/mouuff/mtranslate
"""""""""

import glob
import pandas as pd
from mtranslate import translate
import os
from tqdm import tqdm

#language_code = pd.read_excel(os.getcwd() + "//language_code.xlsx") #If you're working on PC
language_code = pd.read_excel(os.getcwd() + r"/language_code.xlsx") #If you're working on Mac

chosen_lang = input("Which language would you like the file translated into?\n")
while str(language_code[language_code['Language'] == chosen_lang.title()].values) == "[]":
    chosen_lang = input("The language you requested could not be found. Please try again\n")
else:
    chosen_code = language_code[language_code['Language'] == chosen_lang.title()]['Code'].values[0]


def tran(x):
    if str(x) == "nan":
        return ""
    elif isinstance(x, int):
        return ""
    elif isinstance(x, float):
        return ""
    else:
        return translate(str(x), chosen_code, "auto")

# files = glob.glob(os.getcwd() + "//Input/*.xlsx") #If you're working on PC
files = glob.glob(os.getcwd() + r"/Input/*.xlsx") #If you're working on Mac

for fle in files:
    print("\n" + "Working on " + str(fle[fle.find("Input")+6:]))
    df = pd.read_excel(fle)
    out_df = pd.DataFrame()
    for i in tqdm(range (0,len(df))):
        row = df[df.index == i]
        line_to_trans = str(row[row.columns[1]].values[0]).title()
        id = row[row.columns[0]].values[0]
        try:
            tran_sent = tran(line_to_trans)
            if line_to_trans == "Nan":
                print("\n" + "index num " + str(i) + " is empty")
                line_to_trans = ""
                tran_sent = ""
                row_out = pd.DataFrame(data={"ID" : [id], "line_to_translate" : [line_to_trans], "translate_line" : [tran_sent]})
            else:
                #print("index num " + str(i) + " checked/translated")
                row_out = pd.DataFrame(data={"ID" : [id], "line_to_translate" : [line_to_trans], "translate_line" : [tran_sent]})
                out_df = pd.concat([out_df, row_out])
        except:
            print("\n" + "index num " + str(i) + " had a problem")


    #out_df.to_excel(os.getcwd() + "Output\\" + fle[fle.find("Input")+6:], index=False) #If you're working on PC
    out_df.to_excel(os.getcwd() + r"/Output/" + fle[fle.find("Input")+6:], index=False) #If you're working on Mac
