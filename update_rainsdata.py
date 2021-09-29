import sys
import time # 今回は、プログラムをsleepするために使用
import os
from urllib.request import urlopen
from bs4 import BeautifulSoup
import requests
import urllib
import setting as s


# ---------------------------------------------------------------------------------------------
#関数
# ---------------------------------------------------------------------------------------------
#万円削除
def remove_price(price):
    new_p = price.rstrip("万円")
    new_p = new_p.replace(',','')
    return float(new_p)

#坪に変換
def m_tubo(area,kind, r = 0):
    if kind != s.kind[2] or r == 0:
        new_a = area.rstrip("m²")
        new_a = new_a.rstrip("㎡")
        new_a = new_a.replace(',','')
    else:
        new_a = area.split("m²")[0]
        new_a = new_a.replace(',','')
    t = float(new_a)*0.3025
    return t,float(new_a)

#坪単価計算
def t_p(price,area):
    t_p = round(float(price)/area,2)
    return float(t_p)

#万円、坪、坪単価処理まとめ
def str_modify(areas_name,list_name,kind):
    #坪に変換
    tubo = m_tubo(list_name[areas_name],kind,1)
    #価格の万円削除
    new_price = remove_price(list_name["価格"])
    #坪単価計算
    tubo_price = t_p(new_price,tubo[0])
    return new_price,tubo,tubo_price

#数字→アルファベット
def num2alpha(num):
    if num<=26:
        return chr(64+num)
    elif num%26==0:
        return num2alpha(num//26-1)+chr(90)
    else:
        return num2alpha(num//26)+chr(64+num%26)

#スプレッドシートのレンジ指定
def uprange(starta,startn,add_list,endn):
    sa = starta
    sn = startn
    ea = num2alpha(len(add_list))
    en = endn +7
    add_range = sa + str(sn) + ":" + ea + str(en)
    return add_range

#urlを書き込むセル
def link_cell(row,col):
    vrow = row + 2
    vcol = num2alpha(len(col) + 1)
    vc = vcol + str(vrow)
    return vc

#平均値計算
def average(a_list):
    ave = round(sum(a_list) / len(a_list) , 2)
    return ave

#html解析
def find(list_name,kind):
    items = [[i[0],"#" + r_id['id'] + i[1]] for i in list_name[kind]]
    items = {i[0]:soup.select(i[1]) for i in items }
    items = {i[0]:[t.text for t in items[i[0]]] for i in list_name[kind]}
    return items

#坪単価の比較、色つけ
def color(sheet):
    if past_list["坪単価"] < average(r_lists["坪単価"]):
        print(k +": id:" + mid + "の坪単価は過去の取引事例より安いです。")
        for i in sheet:
            i[0].format(i[1], {
                'backgroundColor': {
                    "red": 0.9,
                    "green": 0.9,
                    "blue": 0.0
                }
            }) 

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

print("ファイル名「"+ s.fname + "」に過去の取引事例を書き込みます。")

# ---------------------------------------------------------------------------------------------
#グローバル変数
# ---------------------------------------------------------------------------------------------
past = s.gc.open(s.fname) #物件一覧ファイルを開く
ws_list = [i.title for i in past.worksheets()] #ファイル内のワークシートのリスト一覧
ws_kinds = {i:past.worksheet(i) for i in s.kind} #ファイル内の該当シート[中古マンション][中古戸建て][土地]
num_col_lists = {i:[t for num,t in enumerate(ws_kinds[i].col_values(1)) if num != 0] for i in s.kind} #該当シートのid一覧[中古マンション][中古戸建て][土地]
# ---------------------------------------------------------------------------------------------
#グローバル変数ここまで
# ---------------------------------------------------------------------------------------------


# ---------------------------------------------------------------------------------------------
#取引事例書き込み
# ---------------------------------------------------------------------------------------------
s.times(30)
for k in s.kind:
    print(k + "の取引事例書き込みを開始します。")
    kind = k
    head = s.detail_head[kind]

    #アットホームのスクレイピング結果の数だけ繰り返し
    for mcount,mid in enumerate(num_col_lists[kind]):
        s.times(5)
        past_item = ws_kinds[kind].row_values(mcount + 2) #本物件の情報取得
        past_list = {i:past_item[num] for num,i in enumerate(s.head[kind])} #本物件の情報の辞書化
        sc_path = s.path_route + str(mid) + ".html" #スクレイピングするファイルのパス
        vc = link_cell(mcount,s.head[kind]) #urlを書き込むセル
        
        #スクレイピング未処理であればスクレイピング開始
        if past_list["id"] not in ws_list:
            print(k +": id:" + mid + "の書き込みを開始します。")

            #ローカルフォルダにreinsのhtmlファイルがあれば実行
            if os.path.exists(sc_path):
                print(k +": id:" + mid + "のhtmlファイルが見つかりました。")
                soup = BeautifulSoup(open(sc_path),'html.parser')
                r_id = soup.select_one(s.reins_pid)

                #htmlファイルが該当のページかのチェック
                if r_id:
                    print(k +": id:" + mid + "のhtmlファイルから解析タグが見つかりました。解析を開始します。")
                    id_c = soup.select('#' + r_id['id'] + s.reins_selector_lists[kind][0][1])

                    #reinsの過去の取引事例が1件以上なら実行
                    if len(id_c) > 0:
                        mcom = past.add_worksheet(title=mid,rows=200,cols=20) #ファイルにワークシートを新規作成
                        wbid = past.id #ファイルのID
                        wsid = mcom.id #ワークシートのID
                        ws_url = s.wb_path + str(wbid) + "/edit#gid=" + str(wsid) #新規ワークシートのパス
                        r_lists = find(s.reins_selector_lists,kind) #reinsのスクレイピング結果
                        
                        #中古マンションの場合
                        if k == s.kind[0]:
                            area_name = "専有面積"
                            mage = past_list["築年月"].split("年")[0] #本物件の築年月
                            up_lists = [c for c,i in enumerate(r_lists["築年月"]) if i.split("年")[0] == mage]
                            up_num  = len(up_lists)
                        else:
                            area_name = "土地面積"
                            up_num  = len(id_c)
                            up_lists = list(range(up_num))
                        
                        if up_num > 0:
                            tubo_tanka = str_modify(area_name,past_list,kind) # 該当物件の面積、価格、坪単価の調整
                            past_list["坪単価"] = tubo_tanka[2]
                            new_p = [remove_price(i) for i in r_lists["価格"]] #文字列の調整（価格）
                            new_a = [[m_tubo(i,kind)[p] for i in r_lists[area_name]] for p in range(2)] #文字列の調整（面積）
                            r_lists["坪単価"] = [t_p(i,t) for i,t in zip(new_p,new_a[0])] #坪単価計算→スクレイピング結果に追加

                            # スプレッドシート書き込みのための設定
                            up_cells = uprange("A",1,head,up_num)
                            set_col = len(head)
                            minfo = mcom.range(up_cells)
                            av_cell = set_col * (up_num + 6)
                            area_cell = [c for c,i in enumerate(head) if i == area_name]
                            tp_cell = [c for c,i in enumerate(head) if i == "坪単価"]
                            color_col = num2alpha(tp_cell[0] + 1)
                            color_cell = [[mcom,color_col + "5"],[ws_kinds[kind],vc]]
                            # 該当物件基本情報
                            for i in range(3):
                                minfo[i].value = s.info_head[i]
                                minfo[i+set_col].value = past_list[s.info_head[i]]
                            # 該当物件詳細情報
                            for num, i in enumerate(head):
                                minfo[num + set_col * 3].value = i
                                if i in past_list.keys():
                                    minfo[num + set_col * 4].value = past_list[i]
                            # reinsスクレイピング結果
                            for num,i in enumerate(up_lists):
                                for c,t in enumerate(head):
                                    minfo[c + set_col *(6 + num)].value = r_lists[t][i]

                            # 中古マンションの場合平均値リスト
                            if k == s.kind[0]:
                                new_p = [new_p[i] for i in up_lists]
                                new_a[1] = [new_a[1][i] for i in up_lists]
                                r_lists["坪単価"] = [r_lists["坪単価"][i] for i in up_lists]

                            # 平均値計算
                            minfo[av_cell].value = "平均"
                            minfo[av_cell + 1].value = average(new_p)
                            minfo[av_cell + area_cell[0]].value = average(new_a[1])
                            minfo[av_cell + tp_cell[0]].value = average(r_lists["坪単価"])
                            # 書き込み
                            mcom.update_cells(minfo)
                            print(k +": id:" + mid + "のシートに取引事例を書き込みました。")
                            # タブのリンク貼り付け
                            ws_kinds[kind].update_acell(vc,ws_url)
                            # 平均値比較
                            color(color_cell)
                            print(k +": id:" + mid + "終了")
                        else:
                            ws_kinds[kind].update_acell(vc,"該当件数0件")
                            print(k +": id:" + mid + "の過去の取引事例は有りませんでした。")
                    else:
                        ws_kinds[kind].update_acell(vc,"該当件数0件")
                        print(k +": id:" + mid + "の過去の取引事例は有りませんでした。")
                else:
                    ws_kinds[kind].update_acell(vc,"ファイルが違います。reinsから再ダウンロードしてください。")
                    print(k +": id:" + mid + "のhtmlファイルから解析タグが見つかりませんでした。reinsから再ダウンロードしてください。")
                os.remove(sc_path)
            else:
                ws_kinds[kind].update_acell(vc,"ファイルが見つかりません。reinsから再ダウンロードしてください。")
                print(k +": id:" + mid + "のhtmlファイルが見つかりませんでした。reinsからダウンロードしてください。")
        else:
            print(k +": id:" + mid + "の取引事例は書き込み済みです。")
    s.times(30)
# ---------------------------------------------------------------------------------------------
#取引事例書き込みここまで
# ---------------------------------------------------------------------------------------------
# ---------------------------------------------------------------------------------------------
# 
# 
#処理終了
# 
# 
# ---------------------------------------------------------------------------------------------
