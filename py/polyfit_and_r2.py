import PySimpleGUI as sg
import numpy
import openpyxl
import matplotlib.pyplot as plt
import pandas as pd

sg.theme("GreenMono")
frame1 = sg.Frame('',
    [
        [
            sg.Text('処理を実行したいエクセルファイルを選択してください')
        ],
        [
            sg.Text("ファイル"),
            sg.InputText(key='pfs',enable_events=True),
            sg.FileBrowse('ファイルを選択',target='pfs',file_types=(('Excell ファイル', '*.xlsx'),))
        ],
        [sg.Submit("実行"), sg.Cancel("キャンセル")],
    ], size=(550, 700)
)

frame2 = sg.Frame('',
    [
        [
            sg.Text('実行結果')
        ],
    ] , size=(600, 700)
)

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