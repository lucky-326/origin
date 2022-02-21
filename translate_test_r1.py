#!/usr/bin/env python
# coding: utf-8

# In[1]:


from googletrans import Translator
import tkinter
import tkinter.filedialog as fileDialog
from tkinter import messagebox
import openpyxl

#-----翻訳したいファイルのパス・名前を取得--------
root = tkinter.Tk()                   #ウインドウを生成
root.attributes('-topmost', True)     #topmost指定(最前面に配置)
root.withdraw()                       #空のルートウインドウを非表示
root.focus_force()                    #ウインドウにフォーカスを当てる

#ファイルの拡張子を指定
fileTypes = [('xlsxファイル', '*.xlsx')]
#ダイアログを開く箇所のパスを指定
initialDir = 'C:/Users/lucky/Desktop/Workspace/translate_test_01/'
#ダイアログを開いてファイルのパス・名前を取得
files = fileDialog.askopenfilenames(filetypes = fileTypes, initialdir = initialDir)

#-----翻訳処理----------------------------------

tr = Translator()                            #Googletransライブラリ　  
wb = openpyxl.load_workbook(files[0])        #翻訳したいファイルを読み取り
ws = wb['Sheet1']                            #ワークシートを指定

#1行ずつループして行を取得
for row in ws.iter_rows(min_row=2):          #行の範囲を"row"に格納(min_rowでHeader部分を除いた2行目から指定)
    no = row[0]                              #"No"列を指定
    ja = row[1]                              #"Japanese"列を指定
    zh = row[2]                              #"Chinese"列を指定
    #日本語(ja)から中国語(zh-ch)に翻訳
    trans_zh = tr.translate(ja.value, src="ja", dest="zh-cn").text      
    zh.value = str(trans_zh)                 #"Chinese"列に翻訳したデータをinput
    
wb.save(files[0])                             #翻訳したファイルを上書き保存
#翻訳完了したらメッセージボックスを表示
messagebox.showinfo('翻訳機能', '翻訳完了')

#-----改良予定----------------------------------
#1.例外処理
#2.多言語対応
#3.exe化


# In[ ]:




