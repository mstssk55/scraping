import AppKit
import pyautogui as pg
import pyperclip
import time # 今回は、プログラムをsleepするために使用
import webbrowser
import sys
import getpass # 実行ユーザーを取得するために使用
import requests
import json
import shutil
import os
import glob
import setting as s


# ---------------------------------------------------------------------------------------------
#関数
# ---------------------------------------------------------------------------------------------

#id,住所の取得
def i_a(c):
    col = past_wk.col_values(c)
    del col[0]
    return col
#住所分割処理のパーツ
def add_split(address,key):
    cityname = address.split(key)[0] + key
    addname = address.split(key)[1]
    return cityname,addname

#住所分割処理
def change_address(a_id):
    address_list = []
    if a_id[0] != "余":
        if "市" in a_id:
            address_list = list(add_split(a_id,"市"))
            if address_list[0] == "札幌市":
                word = address_list[1].split("区")[0]
                city = address_list[0] + word + "区"
                add = address_list[1].split("区")[1]
                address_list = [city,add]
        elif "町" in a_id:
            address_list = list(add_split(a_id,"町"))
        elif "村" in a_id:
            address_list = list(add_split(a_id,"村"))
    else:
        if "町" in a_id:
            address_list = list(add_split(a_id,"町"))
        elif "村" in a_id:
            address_list = list(add_split(a_id,"村"))
    return address_list

#pyautoguiクリック
def pgclick(pl,c = 1):
    pg.click(
        x=pl[0], 
        y=pl[1], 
        clicks=c,
        interval=c - 1, 
        button='left'
    )
    time.sleep(0.3)

def pgpress(key):
    pg.press(key)
    time.sleep(0.5)

#pyautogui hotkey
def hot(k1,k2):
    pg.keyDown(k1)
    time.sleep(0.2)
    pg.keyDown(k2)
    time.sleep(0.2)
    pg.keyUp(k2)
    time.sleep(0.2)
    pg.keyUp(k1)
    time.sleep(0.2)

#pyautogui ペースト
def paste(address):
    pyperclip.copy(address)
    hot("command","v")

#pyautogui ２回目以降の住所入力
def paste2(p,address):
    pgclick(p)
    hot("command","a")
    pg.press('delete')
    time.sleep(0.2)
    pyperclip.copy(address)
    hot("command","v")

#ページ遷移のチェック
def c_url(p,n):
    pgclick(p,2)
    time.sleep(3)
    hot("command","l")
    hot("command","c")
    aurl = pyperclip.paste()
    if n == 0:
        if aurl != s.s_url:
            print("検索ページに遷移していません。売買検索をクリックしてください。")
            input("ページ遷移後にターミナル上でenterを押してください。: ")
            pgclick(s.point_lists[0]) #下部のダウンロード削除
    elif n == 1:
        if aurl != s.r_url:
            print("検索一覧ページに遷移していません。住所などが正しく入力されているか確認してください。")
            input("修正後に検索ボタンを押して、ページ遷移後にターミナル上でenterを押してください。: ")
            pgclick(s.point_lists[0]) #下部のダウンロード削除
    print("正常にページが遷移しました。")


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

print("ファイル名「"+ s.fname + "」からreinsの過去の取引事例を検索します")

# ---------------------------------------------------------------------------------------------
#html取得の事前準備
# ---------------------------------------------------------------------------------------------

#htmlを取得するファイルを開く
past = s.gc.open(s.fname)
#ローカルのファイルリストを取得
l_files = glob.glob(s.path_route + "*.html")
#ローカルのファイルリストを名前だけに変換
lf_list = [os.path.split(file)[1].split(".")[0] for file in l_files]
print("ローカルファイルの名前一覧を取得しました")
# ---------------------------------------------------------------------------------------------
#html取得の事前準備ここまで
# ---------------------------------------------------------------------------------------------

# ---------------------------------------------------------------------------------------------
#reins自動操作開始
# ---------------------------------------------------------------------------------------------


for k in s.kind:
    kind = k
    past_wk = past.worksheet(kind)
    #ファイルからid列と住所列を取得してリストに格納
    num_col = i_a(1)
    add_col = i_a(5)
    add_list = {num_col[num]:i for num,i in enumerate(add_col)}
    #ローカルファイルと取得ファイルの重複チェック
    compare = set(lf_list) & set(num_col)
    s_list = compare ^ set(num_col)
    print(s.fname + "の" + kind + "は" + str(len(add_list)) + "件です")
    #自動操作するリスト
    s_list = list(s_list)
    if len(s_list) > 0:
        print(str(len(add_list)) + "件中" +str(len(compare)) + "件は取得済みです")
        print(str(len(s_list)) + "件の自動操作を開始します")

        #reinsページ
        webbrowser.open(s.t_url)
        print("reinsページを開きました")
        time.sleep(3)
        #画面サイズ調整
        hot("command","l")
        hot("command","0")
        hot("command","-")
        hot("command","-")
        hot("command","-")

        pgclick(s.point_lists[0])
        print("画像サイズを調整しました。")

        #自動操作するリストの数だけ繰り返し処理
        for count,mid in enumerate(s_list):
            print(kind +":" + str(len(s_list)) + "件中" + str(count + 1) + "件目の処理を開始します")
            print("ID：" + mid)
            #住所の分割
            city = change_address(add_list[mid])[0]
            add = change_address(add_list[mid])[1]
            print("【市区】：" + city + "　【丁目】：" + add)
            if count == 0: #一回目は対象区分等の選択
                #topページ-------------------------------------------------------------
                
                c_url(s.point_lists[1],0) # 売買物件検索クリック
                #検索ページ-------------------------------------------------------------
                pgclick(s.point_lists[2]) #対象区分成約クリック
                pgclick(s.point_lists[3]) #物件種別セレクトＢＯＸクリック
                pgclick(s.b_kind[kind]) #物件種別クリック
                pgclick(s.point_lists[4],2) #都道府県名テキストボックスクリック
                paste("北海道") #都道府県名入力
                pgclick(s.point_lists[5]) #所在地名１テキストボックスクリック
                paste(city) #所在地名１入力
                pgclick(s.point_lists[6]) #所在地名２テキストボックスクリック
                paste(add) #所在地名２入力

                pgclick(s.point_lists[10]) #あいてる場所クリック
                hot("command","down")
                time.sleep(1)
                pgclick(s.point_lists[9]) #日付を指定クリック
                print(s.point_lists[9])
                pgpress('tab')
                pgpress('down')
                pgpress('down')
                pgpress('enter')
                pgpress('tab')
                pgpress('1')
                pgpress('tab')
                pgpress('1')
                pgpress('tab')
                pgpress('8')
                print("検索条件指定完了")
            else: #二回目以降は住所の変更のみ
                #検索ページ-------------------------------------------------------------
                paste2(s.point_lists[5],city) #所在地名１テキストボックスの処理
                paste2(s.point_lists[6],add) #所在地名２テキストボックスの処理
                time.sleep(1)
                print("検索条件指定完了")
            c_url(s.point_lists[7],1) #検索ボタンクリック
            #検索結果ページ-------------------------------------------------------------
            num = 0
            while not os.path.exists(s.path_route + str(mid) + ".html"): #htmlダウンロード
                if num >2: #3回以上エラーが出れば一時停止
                    input("一度処理を停止して再度実行してください。: ")
                    pgclick(s.point_lists[0]) #下部のダウンロード削除
                pg.press('esc')
                pg.press('esc')
                hot("command","s")
                hot("command","a") 
                pg.press('delete')
                time.sleep(0.2)
                dfname = mid + '.html'
                paste(dfname)
                pg.press('enter')
                time.sleep(2)
                num +=1
                if os.path.exists(s.path_route + str(mid) + ".html"):
                    shutil.rmtree(s.path_route + mid + '_files') #必要のないファイルを削除
                    print(mid + "の該当ファイルが正常にダウンロードされました")
            pgclick(s.point_lists[0]) #下部のダウンロード削除
            pgclick(s.point_lists[8]) #前のページへ戻る
            if count == 0:
                pgclick(s.point_lists[10]) #あいてる場所クリック
                hot("command","up")

            print(kind +":" + str(len(s_list)) + "件中" + str(count + 1) + "件目終了")
        print(kind +":すべてのhtmlファイルを取得しました。（" + str(len(s_list)) + "件）" )
        hot("command","w")
    else:
        print(k + "のhtmlファイルはすべて取得済みです。")


print("すべての処理が終了しました")

# ---------------------------------------------------------------------------------------------
# 
# 
#処理終了
# 
# 
# ---------------------------------------------------------------------------------------------
