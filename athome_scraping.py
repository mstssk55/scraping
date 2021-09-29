# -*- coding: utf-8 -*-
from selenium import webdriver # installしたseleniumからwebdriverを呼び出せるようにする
from selenium.webdriver.common.keys import Keys # webdriverからスクレイピングで使用するキーを使えるようにする。
from selenium.webdriver.support.select import Select
from selenium.webdriver.chrome.options import Options
import requests
import sys
import setting as s


# ---------------------------------------------------------------------------------------------
#関数
# ---------------------------------------------------------------------------------------------

#アットホームのセレクトボックスを新着順に並べ替え
def selectbox():
    s.times(5)
    # ドロップダウンから新着順を選択
    dropdown=driver.find_element_by_class_name("bukken-sort")#セレクトボックスのクラス名
    select = Select(dropdown)
    select.select_by_value('33')#新着順のvalue
    s.times(5)

#スクレイピング結果をスプレッドシートに書き込み
def update(k):
    sapporo_num = [c for c,i in enumerate(property_items["所在地"]) if "札幌市" in i and property_items["id"][c] not in past_data[k]]
    ds = sheet_list[k].range('A2:L100')
    dl = len(past_data[k])
    c_num = "A" + str(dl + 2)
    h_num = "L" + str(dl + 100)
    past_range = c_num + ":" + h_num
    pu = past_sheet[k].range(past_range)
    for num, i in enumerate(sapporo_num):
        for c,t in enumerate(s.head[k]):
            ds[c+num*12].value = property_items[t][i]
            pu[len(s.head[k])+num*12].value = s.today
            pu[c+num*12].value = property_items[t][i]
    sheet_list[k].update_cells(ds)
    print("ファイル名「" + s.fname + "」に" + k + "のスクレイピング結果を書き込みました。")
    past_sheet[k].update_cells(pu)
    print("ファイル名「" + s.pfname + "」に" + k + "のスクレイピング結果を書き込みました。")

def find(list,kind):
    items = [[i[0],eval(i[2])(i[1]),i[3]] for i in list[kind]]
    items = {i[0]:[t.get_attribute(i[2]) if i[2] != "text" else t.text for t in i[1]] for i in items}
    return items

# ---------------------------------------------------------------------------------------------
#関数ここまで
# ---------------------------------------------------------------------------------------------


# ---------------------------------------------------------------------------------------------
# 
# 
#処理開始
# 
# 
# ---------------------------------------------------------------------------------------------
print("アットホームのスクレイピング処理を開始します。")

# ---------------------------------------------------------------------------------------------
#スクレイピング事前準備
# ---------------------------------------------------------------------------------------------

#driveフォルダ内のファイル名一覧
f_list = [f['title'] for f in s.drive.ListFile({'q': '"{}" in parents'.format(s.folder_id)}).GetList()]
print("フォルダ内のファイル一覧を取得しました。")

#フォルダ内に新規ファイル作成
if s.fname not in f_list: #ファイルが作成されていなければファイルを作成する
    f = s.drive.CreateFile({
        'title': s.fname,
        'mimeType': 'application/vnd.google-apps.spreadsheet',
        "parents": [{"id": s.folder_id}]})
    f.Upload()
    print("ファイル名:「" + s.fname + "」を作成しました。")
    wb = s.gc.open_by_key(f['id']) #作成したファイルを開く
    wkm = wb.sheet1
    wkm.update_title(s.kind[0])
    sheet_list = {i if num !=0 else i :wb.add_worksheet(title=i,rows=200,cols=20)  if num !=0 else wkm for num,i in enumerate(s.kind)}
    print("ファイル名:「" + s.fname + "」のシート名を変更しました。")

    # 1行目の書き込み
    for count, t in enumerate(sheet_list):
        hk = s.kind[count]
        uc = sheet_list[hk].range('A1:T1')
        for num, i in enumerate(s.head[hk]):
            uc[num].value = i
        sheet_list[hk].update_cells(uc)
    print("ファイル名:「" + s.fname + "」の各シートにヘッダー行を入力しました。")

else: #すでにファイルが作成されていればシートを変数に格納
    print("ファイル名:「" + s.fname + "」はすでに作成済みです。")
    wb = s.gc.open(s.fname) #作成済みファイルを開く
    sheet_list = {i:wb.worksheet(i) for i in s.kind}

#種類ごとにスクレイピング済みかのチェック（id列の数を判定。1ならスクレイピング実行）
id_nums = {i:len(sheet_list[i].col_values(1)) for i in sheet_list}
#過去の物件一覧取得（重複の判定）
past = s.gc.open(s.pfname) #物件一覧ファイルを開く
#過去のスクレイピング済みリスト
past_data = {i:[t for num,t in enumerate(past.worksheet(i).col_values(1)) if num !=0] for i in s.kind}
#過去のスクレイピングファイルのシート名
past_sheet = {i:past.worksheet(i) for i in s.kind}
# ---------------------------------------------------------------------------------------------
#スクレイピング事前準備ここまで
# ---------------------------------------------------------------------------------------------


# ---------------------------------------------------------------------------------------------
#アットホームスクレイピング処理
# ---------------------------------------------------------------------------------------------

for k in s.kind:
    if id_nums[k] == 1:
        print(k + "のスクレイピングを実行します。")
        driver = webdriver.Chrome(s.c_path + '/chromedriver') 
        driver.get(s.spages[k]) #スクレイピングするページ
        if len(driver.find_elements_by_class_name("bukken-sort")) > 0:
            print(k + "の該当ページが正常に開けました。")
            selectbox()
            print("新着順に並べ替えました")
            property_items = find(s.ah_selector_lists,k)
            print("アットホームから" + k + "の情報を取得しました。")
            driver.quit()
            update(k)
            s.times(2)
            print(k + "のスクレイピングが終了しました。")
        else:
            print(k + "の該当ページが正常に開けませんでした。" + k + "のスクレイピング処理をスキップします。")
    else:
        print(k + "のスクレイピングは実行済みです。")
# ---------------------------------------------------------------------------------------------
#アットホームスクレイピング処理ここまで
# ---------------------------------------------------------------------------------------------

print("本日のスクレイピングは終了しました。")

# ---------------------------------------------------------------------------------------------
# 
# 
#処理終了
# 
# 
# ---------------------------------------------------------------------------------------------
