import datetime
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from pydrive.auth import GoogleAuth
from pydrive.drive import GoogleDrive
import sys
import time # 今回は、プログラムをsleepするために使用
import configparser
import os
import json
import getpass

c_path = os.path.dirname(os.path.abspath(sys.argv[0]))
config = configparser.ConfigParser()
config.read(c_path + '/config.ini', encoding='utf-8')
# config.read('/Users/masato/Desktop/athome_scraping/dist/config.ini', encoding='utf-8')

# -----------------------------
# モード設定
# テストモード=0
# 本番=1
# -----------------------------


#変更必要--------------------------------------------------------------------------------------------
mode = int(config['MODE']['Mode_Num'])
test_file_name = config['MODE']['Test_File_Name']
path_route = config['USER']['DOWNLOAD_PATH']
#変更必要--------------------------------------------------------------------------------------------

def p(t):
    print(t)
    sys.exit()

def times(t):
    time.sleep(t)

def conf(h,d):
    conf_data = json.loads(config.get(h,d))
    return conf_data

# ------------------------------------------------------------------------

# 共通設定

# ------------------------------------------------------------------------

# -----------------------------
# 基本設定
# -----------------------------
#種類
kind = [
    "中古マンション",
    "中古戸建て",
    "土地"
]

# 日付取得
today = str(datetime.date.today()) 

if mode == 0:
    #ファイル名
    fname = test_file_name
    #過去の物件一覧ファイル名
    pfname = "st"
    #スクレイピングするページ
elif mode ==1:
    #ファイル名
    fname = today
    #過去の物件一覧ファイル名
    pfname = "物件一覧"
    #スクレイピングするページ


spages={
    kind[0]:"https://www.athome.co.jp/mansion/chuko/hokkaido/list/",
    kind[1]:"https://www.athome.co.jp/kodate/hokkaido/list/",
    kind[2]:"https://www.athome.co.jp/tochi/hokkaido/list/"
}

# -----------------------------
# googledrive.spreadsheetの設定
# -----------------------------

scope = ['https://spreadsheets.google.com/feeds',
         'https://www.googleapis.com/auth/drive']
credentials = ServiceAccountCredentials.from_json_keyfile_name(c_path + '/js/athome-309806-6ba0b65f0ca5.json', scope)
gauth = GoogleAuth()
gauth.credentials = credentials
drive = GoogleDrive(gauth)
folder_id = 'folder_id' #google driveフォルダーのid
gc = gspread.authorize(credentials)


# ------------------------------------------------------------------------

# アットホームスクレイピング

# ------------------------------------------------------------------------

# スプレッドシート1行目
head = {
    kind[0]:[
        "id",
        "物件名",
        "url",
        "価格",
        "所在地",
        "交通",
        "階建",
        "間取り",
        "専有面積",
        "築年月",
        "構造"
    ],
    kind[1]:[
        "id",
        "物件名",
        "url",
        "価格",
        "所在地",
        "交通",
        "間取り",
        "建物面積",
        "土地面積",
        "築年月",
    ],
    kind[2]:[
        "id",
        "物件名",
        "url",
        "価格",
        "所在地",
        "交通",
        "土地面積",
        "私道負担面積",
        "権利",
        "最適用途",
        "建ぺい率/容積率"
    ]
}

# スクレイピングするエレメントの検索方法
#0=class
#1=xpath
findby = [
    "driver.find_elements_by_css_selector",
    "driver.find_elements_by_xpath"
]
# スクレイピングするエレメントの検索方法
#0=get_attribute('data-bukken-no')
#1=get_attribute('href')
#2=text
text_selector = [
    'data-bukken-no',
    'href',
    "text"
]

ah_selector = {
    kind[0]:[
        ['.object.boxHover',0,0], #物件ID
        ['.kslisttitle.boxHoverLinkStop',0,2], #物件名
        ['.kslisttitle.boxHoverLinkStop',0,1], #URL
        ['.fwB.fcRed.fs16',0,2], #価格
        ['//*[@id="item-list"]/div/div/div[2]/table/tbody/tr[2]/td',1,2], #所在地
        ['//*[@id="item-list"]/div/div/div[2]/table/tbody/tr[3]/td',1,2], #交通
        ['//*[@id="item-list"]/div/div/div[2]/table/tbody/tr[4]/td',1,2], #階建
        ['//*[@id="item-list"]/div/div/div[2]/table/tbody/tr[5]/td[1]',1,2], #間取り
        ['//*[@id="item-list"]/div/div/div[2]/table/tbody/tr[5]/td[2]',1,2], #専有面積
        ['//*[@id="item-list"]/div/div/div[2]/table/tbody/tr[6]/td[1]',1,2], #築年月
        ['//*[@id="item-list"]/div/div/div[2]/table/tbody/tr[6]/td[2]',1,2] #構造
    ],
    kind[1]:[
        ['.object.boxHover',0,0], #物件ID
        ['.kslisttitle.boxHoverLinkStop',0,2], #物件名
        ['.kslisttitle.boxHoverLinkStop',0,1], #URL
        ['.fwB.fcRed.fs16',0,2], #価格
        ['//*[@id="item-list"]/div/div/div[2]/table/tbody/tr[2]/td',1,2], #所在地
        ['//*[@id="item-list"]/div/div/div[2]/table/tbody/tr[3]/td',1,2], #交通
        ['//*[@id="item-list"]/div/div/div[2]/table/tbody/tr[4]/td',1,2], #間取り
        ['//*[@id="item-list"]/div/div/div[2]/table/tbody/tr[5]/td',1,2], #建物面積
        ['//*[@id="item-list"]/div/div/div[2]/table/tbody/tr[6]/td',1,2], #土地面積
        ['//*[@id="item-list"]/div/div/div[2]/table/tbody/tr[7]/td',1,2] #築年月
    ],
    kind[2]:[
        ['.object.boxHover',0,0], #物件ID
        ['.kslisttitle.boxHoverLinkStop',0,2], #物件名
        ['.kslisttitle.boxHoverLinkStop',0,1], #URL
        ['.fwB.fcRed.fs16',0,2], #価格
        ['//*[@id="item-list"]/div[1]/div/div/div[2]/table/tbody/tr[3]/td',1,2], #所在地
        ['//*[@id="item-list"]/div[1]/div/div/div[2]/table/tbody/tr[4]/td',1,2], #交通
        ['//*[@id="item-list"]/div[1]/div/div/div[2]/table/tbody/tr[5]/td',1,2], #土地面積
        ['//*[@id="item-list"]/div[1]/div/div/div[2]/table/tbody/tr[6]/td',1,2], #私道負担面積
        ['//*[@id="item-list"]/div[1]/div/div/div[2]/table/tbody/tr[7]/td',1,2], #権利
        ['//*[@id="item-list"]/div[1]/div/div/div[2]/table/tbody/tr[8]/td',1,2], #最適用途
        ['//*[@id="item-list"]/div[1]/div/div/div[2]/table/tbody/tr[9]/td',1,2] #建ぺい率／容積率
    ]
}
ah_selector_lists = {i:[[h,s[0],findby[s[1]],text_selector[s[2]]] for h,s in zip(head[i],ah_selector[i])] for i in kind}


# ------------------------------------------------------------------------

# レインズ自動操作

# ------------------------------------------------------------------------

#レインズページURL
t_url = config['URL']['Top_Url']#rainsトップページ
s_url = config['URL']['Search_Url']#検索条件入力ページ
r_url = config['URL']['Result_Url']#検索結果ページ

#レインズページクリック座標
point_lists = [
    conf('CLICK','Click_01'),
    conf('CLICK','Click_02'),
    conf('CLICK','Click_03'),
    conf('CLICK','Click_04'),
    conf('CLICK','Click_05'),
    conf('CLICK','Click_06'),
    conf('CLICK','Click_07'),
    conf('CLICK','Click_08'),
    conf('CLICK','Click_09'),
    conf('CLICK','Click_10'),
    conf('CLICK','Click_11')
]
#レインズページクリック座標（物件種別セレクトボックス）
b_kind={
        kind[0]:conf('SELECT','Select_01'),
        kind[1]:conf('SELECT','Select_02'),
        kind[2]:conf('SELECT','Select_03')
}

# ------------------------------------------------------------------------

# レインズhtmlをスプレッドシートに書き込み

# ------------------------------------------------------------------------

wb_path = "https://docs.google.com/spreadsheets/d/"
reins_pid = '.tab-pane.p-tab.active.card-body'

#基本情報ヘッダー
info_head =[
    "id",
    "物件名",
    "url"
]
#詳細情報ヘッダー
detail_head ={
    kind[0]:[
        "id",
        "価格",
        "所在地",
        "沿線",
        "交通",
        "階建",
        "間取り",
        "専有面積",
        "坪単価",
        "築年月",
        "成約年月日",
        "物件種目",
        "物件名"
    ],
    kind[1]:[
        "id",
        "価格",
        "所在地",
        "沿線",
        "交通",
        "間取り",
        "建物面積",
        "土地面積",
        "坪単価",
        "築年月",
        "成約年月日",
        "物件種目"
    ],
    kind[2]:[
        "id",
        "価格",
        "所在地",
        "沿線",
        "交通",
        "土地面積",
        "坪単価",
        "建ぺい率",
        "容積率",
        "成約年月日",
        "物件種目",
        "用途地域",
        "接道状況",
        "接道"
    ]
}

reins_selector = {
    kind[0]:[
            ' > div > div.p-table.small > div.p-table-body > div> div:nth-child(4)', #id
            ' > div > div.p-table.small > div.p-table-body > div > div.p-table-body-item.font-weight-bold', #価格
            ' > div > div.p-table.small > div.p-table-body > div > div:nth-child(7)', #所在地
            ' > div > div.p-table.small > div.p-table-body > div > div:nth-child(19)', #沿線
            ' > div > div.p-table.small > div.p-table-body > div > div:nth-child(20)', #交通
            ' > div > div.p-table.small > div.p-table-body > div > div:nth-child(13)', #階建
            ' > div > div.p-table.small > div.p-table-body > div > div:nth-child(14)', #間取り
            ' > div > div.p-table.small > div.p-table-body > div > div:nth-child(6)', #専有面積
            '', #坪単価
            ' > div > div.p-table.small > div.p-table-body > div > div:nth-child(23)', #築年月
            ' > div > div.p-table.small > div.p-table-body > div > div:nth-child(21)', #成約年月日
            ' > div > div.p-table.small > div.p-table-body > div > div:nth-child(5)', #物件種目
            ' > div > div.p-table.small > div.p-table-body > div > div:nth-child(12)', #物件名

    ],
    kind[1]:[
            ' > div > div.p-table.small > div.p-table-body > div > div:nth-child(4)', #id
            ' > div > div.p-table.small > div.p-table-body > div > div.p-table-body-item.font-weight-bold', #価格
            ' > div > div.p-table.small > div.p-table-body > div > div:nth-child(7)', #所在地
            ' > div > div.p-table.small > div.p-table-body > div > div:nth-child(16)', #沿線
            ' > div > div.p-table.small > div.p-table-body > div > div:nth-child(17)', #交通
            ' > div > div.p-table.small > div.p-table-body > div > div:nth-child(13)', #間取り
            ' > div > div.p-table.small > div.p-table-body > div > div:nth-child(11)', #建物面積
            ' > div > div.p-table.small > div.p-table-body > div > div:nth-child(6)', #土地面積
            '', #坪単価
            ' > div > div.p-table.small > div.p-table-body > div > div:nth-child(22)', #築年月
            ' > div > div.p-table.small > div.p-table-body > div > div:nth-child(18)', #成約年月日
            ' > div > div.p-table.small > div.p-table-body > div > div:nth-child(5)', #物件種目

    ],
    kind[2]:[
            ' > div > div.p-table.small > div.p-table-body > div > div:nth-child(4)', #id
            ' > div > div.p-table.small > div.p-table-body > div > div.p-table-body-item.font-weight-bold', #価格
            ' > div > div.p-table.small > div.p-table-body > div > div:nth-child(7)', #所在地
            ' > div > div.p-table.small > div.p-table-body > div > div:nth-child(12)', #沿線
            ' > div > div.p-table.small > div.p-table-body > div > div:nth-child(13)', #交通
            ' > div > div.p-table.small > div.p-table-body > div > div:nth-child(6)', #土地面積
            '', #坪単価
            ' > div > div.p-table.small > div.p-table-body > div > div:nth-child(15)', #建ぺい率
            ' > div > div.p-table.small > div.p-table-body > div > div:nth-child(19)', #容積率
            ' > div > div.p-table.small > div.p-table-body > div > div:nth-child(18)', #成約年月日
            ' > div > div.p-table.small > div.p-table-body > div > div:nth-child(5)', #物件種目
            ' > div > div.p-table.small > div.p-table-body > div > div:nth-child(10)', #用途地域
            ' > div > div.p-table.small > div.p-table-body > div > div:nth-child(20)', #接道状況
            ' > div > div.p-table.small > div.p-table-body > div > div:nth-child(23)', #接道

    ]
}

reins_selector_lists = {i:[[h,x] for h,x in zip(detail_head[i],reins_selector[i])] for i in kind}
