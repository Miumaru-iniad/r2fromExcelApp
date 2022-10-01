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
        # x = np.linspace(0, 2*np.pi, 500)
        x = np.arange(0, 2*np.pi, 0.05*np.pi)
        ax.plot(x, np.sin(x))
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
        [
            sg.Text("次数を選択してください"),sg.Combo([1,2,3,4,5,6], default_value="選択して下さい",readonly=True,size=(30,1))
        ],
        [sg.Button('Display',key='-display-'), sg.Button('clear',key='-clear-'), sg.Cancel()],
        [sg.Canvas(key='-CANVAS-')],
        [
            sg.Text("相関係数"),sg.InputText("r2=",readonly=True,key="-r2-")
        ],
        [sg.Button("実行",size=(30,2))],
    ], size=(700, 800)
)

frame2 = sg.Frame('',
    [
        [
            sg.Text('実行結果')
        ],
        [
            sg.Text("近似式")
        ],
        [
            sg.Text("1"),sg.InputText("ここに数式が出ます",readonly=True,key="-f1-")
        ],
        [
            sg.Text("2"),sg.InputText("ここに数式が出ます",readonly=True,key="-f2-")
        ],
        [
            sg.Text("3"),sg.InputText("ここに数式が出ます",readonly=True,key="-f3-")
        ],
        [
            sg.Text("4"),sg.InputText("ここに数式が出ます",readonly=True,key="-f4-")
        ],
        [
            sg.Text("5"),sg.InputText("ここに数式が出ます",readonly=True,key="-f5-")
        ],
        [
            sg.Text("6"),sg.InputText("ここに数式が出ます",readonly=True,key="-f6-")
        ],
        [
            sg.Text("結果をエクセルに記入する")
        ],
        [
            sg.Button("記入",size=(25,1)),sg.Button("終了",size=(25,1))
        ],
    ] , size=(600, 800)
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

    elif event == '-display-':
        fig = make_data_fig(fig, make=True)
        fig_agg.draw()

    elif event == '-clear-':
        fig = make_data_fig(fig, make=False)
        fig_agg.draw()

window.close()