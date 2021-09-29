# -*- coding: utf-8 -*-
import AppKit
import pyautogui as pg
import sys

print('中断するにはCrt+Cを入力してください。')
try:
    while True:
        x=input("取得したい箇所にカーソルを当てEnterキー押してください\n")
        x,y =pg.position()
        print(str(x)+','+str(y))

except KeyboardInterrupt:
    print('\n終了')
    sys.exit()