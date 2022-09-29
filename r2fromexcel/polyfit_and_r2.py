import PySimpleGUI as sg
import numpy
import openpyxl
import matplotlib.pyplot as plt
import pandas as pd

sg.theme("GreenMono")
frame1 = sg.Frame('',[] , size=(200, 400))
frame2 = sg.Frame('',[] , size=(400, 400))
layout = [
    [
        frame1,
        frame2
    ]
]
window = sg.Window('近似関数、相関係数アプリ', layout, resizable=True)

#GUI表示実行部分
while True:
    # ウィンドウ表示
    event, values = window.read()

    #クローズボタンの処理
    if event is None:
        print('exit')
        break

window.close()