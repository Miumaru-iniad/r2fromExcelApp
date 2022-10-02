from fileinput import filename
from tabnanny import filename_only
import PySimpleGUI as sg
import numpy as np
import openpyxl
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import pandas as pd

fig = plt.figure(figsize=(5, 4))
ax = fig.add_subplot(111)

def make_data_fig(fig,make = True):

    if make:
        filename = values["-filename-"]
        df = pd.read_excel(filename,sheet_name=0)
        xs = list(df["x"])
        ys = list(df["y"])
        print(xs,ys)
        ax.scatter(xs, ys)
        return fig

    else:
        ax.cla()
        return fig

def draw_figure(canvas, figure):
    figure_canvas = FigureCanvasTkAgg(figure, canvas)
    figure_canvas.draw()
    figure_canvas.get_tk_widget().pack(side='top', fill='both', expand=1)
    return figure_canvas




sg.theme("GreenMono")
frame1 = sg.Frame('',
    [
        [
            sg.Text('処理を実行したいエクセルファイルを選択してください')
        ],
        [
            sg.Text("ファイル"),
            sg.InputText(key='-filename-',enable_events=True,readonly=True),
            sg.FileBrowse('ファイルを選択',target='-filename-',file_types=(('Excell ファイル', '*.xlsx'),))
        ],
        [sg.Button('Display',key='-display-'), sg.Button('clear',key='-clear-')],
        [sg.Canvas(key='-CANVAS-')],
        [
            sg.Text("相関係数"),sg.InputText("r2=",readonly=True,key="-r2-")
        ],
    ], size=(700, 650)
)

frame2 = sg.Frame('',
    [    
        [
            sg.Text('チェックを入れると近似式とグラフを表示する')
        ],
        [
            sg.Checkbox("D=1", enable_events=True, key="-d1-"),sg.InputText("ここに数式が出ます",readonly=True,key="-f1-")
        ],
        [
            sg.Checkbox("D-2", enable_events=True, key="-d2-"),sg.InputText("ここに数式が出ます",readonly=True,key="-f2-")
        ],
        [
            sg.Checkbox("D-3", enable_events=True, key="-d3-"),sg.InputText("ここに数式が出ます",readonly=True,key="-f3-")
        ],
        [
            sg.Checkbox("D-4", enable_events=True, key="-d4-"),sg.InputText("ここに数式が出ます",readonly=True,key="-f4-")
        ],
        [
            sg.Checkbox("D-5", enable_events=True, key="-d5-"),sg.InputText("ここに数式が出ます",readonly=True,key="-f5-")
        ],
        [
            sg.Checkbox("D-6", enable_events=True, key="-d6-"),sg.InputText("ここに数式が出ます",readonly=True,key="-f6-")
        ],
        [
            sg.Text("結果をエクセルに記入する")
        ],
        [
            sg.Button("記入",size=(25,1)),sg.Button("終了",size=(25,1))
        ],
    ] , size=(520, 650)
)

layout = [
    [
        frame1,
        frame2
    ]
]
window = sg.Window('近似関数、相関係数アプリ', layout, resizable=True,finalize=True)


# figとCanvasを関連付ける
fig_agg = draw_figure(window['-CANVAS-'].TKCanvas, fig)

#GUI表示実行部分
while True:
    # ウィンドウ表示
    event, values = window.read()
    #クローズボタンの処理
    if event in (None,"Cancel","終了"):
        print('exit')
        break

    elif event == '-filename-':
        fig = make_data_fig(fig, make=True)
        fig_agg.draw()


    elif event == '-display-':
        fig = make_data_fig(fig, make=True)
        fig_agg.draw()

    elif event == '-clear-':
        fig = make_data_fig(fig, make=False)
        fig_agg.draw()

window.close()