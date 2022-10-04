from fileinput import filename
from msilib.schema import CheckBox
from tabnanny import filename_only
import PySimpleGUI as sg
import numpy as np 
import openpyxl
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import pandas as pd
import sympy


ds=[]
for i in range(1,10):
    ds.append(["-d"+str(i)+"-",i,"D="+str(i),"-f"+str(i)+"-"])
d_values = []
for d in ds:
    d_values.append(d[0])


x_latent = range(0,101)
txt1="ここに数式が出ます" 
txt2="ファイルを選択してください"
txt3="処理を実行しています"

fig = plt.figure(figsize=(5, 4))
ax = fig.add_subplot(111)

checkbox_v=0


def make_data_fig_selected(fig):
    filename = values["-filename-"]
    df = pd.read_excel(filename,sheet_name=0)
    xs = list(df["x"])
    ys = list(df["y"])
    ax.scatter(xs, ys,label="observed")
    r2 = np.corrcoef(xs,ys)[0][1]
    rr2 = round(r2,4)
    window["-r2-"].Update(rr2)
    ax.legend()
    return fig

def make_data_fig(fig,make = True):
    if make:
        filename = values["-filename-"]
        df = pd.read_excel(filename,sheet_name=0)
        xs = list(df["x"])
        ys = list(df["y"])
        polyfit_lst = []
        for i in range(1,11):
            polyfit_lst.append(np.polyfit(xs,ys,i))
        print(polyfit_lst)



        ax.scatter(xs, ys,label="observed")

        x_latent = np.linspace(min(xs), max(xs), 100)


        for d in ds:
            if values[d[0]]:
                fitted_curve = np.poly1d((lambda x, y: np.polyfit(x, y, d[1]))(xs, ys))(x_latent)
                ax.plot(x_latent, fitted_curve, label=d[2])



        r2 = np.corrcoef(xs,ys)[0][1]
        rr2 = round(r2,4)
        window["-r2-"].Update(rr2)
        ax.legend()
        return fig

    else:
        ax.cla()
        return fig

def selected_file():
    for d in ds:
        window[d[0]].Update(value=False)
        window[d[3]].Update(txt1)

def draw_figure(canvas, figure):
    figure_canvas = FigureCanvasTkAgg(figure, canvas)
    figure_canvas.draw()
    figure_canvas.get_tk_widget().pack(side='top', fill='both', expand=1)
    return figure_canvas

def selected_checbox():
    window[ds[checkbox_v][0]].Update(txt2)
    if values["-filename-"] != "":
        window[ds[checkbox_v][3]].Update(txt3)
        fig = make_data_fig(fig, make=False)
        fig = make_data_fig(fig, make=True)
        fig_agg.draw()
        if values[ds[checkbox_v][0]] == False:
            window[ds[checkbox_v][3]].Update(txt1) 



sg.theme("GreenMono")
frame1 = sg.Frame('',
    [
        [
            sg.Text('処理を実行したいファイルを選択してください')
        ],
        [
            sg.Text("ファイル"),
            sg.InputText(key='-filename-',enable_events=True,readonly=True),
            sg.FileBrowse('ファイルを選択',target='-filename-',file_types=(('Excell ファイル', '*.xlsx'),))
        ],
        [sg.Canvas(key='-CANVAS-')],
        [
            sg.Text("相関係数"),sg.InputText("r2=",readonly=True,key="-r2-")
        ],
    ], size=(700, 630)
)
  
col = [[sg.Text('チェックを入れると近似式とグラフを表示する')]]
        # Create several similar fire buttons in the vertical column
for d in ds:
    col += [[sg.Checkbox(d[2], enable_events=True, key=d[0]),sg.InputText("ここに数式が出ます",readonly=True,key=d[3])]]

frame2 = sg.Frame('',
    [    
        [sg.Column(col)],
        [
            sg.Text("結果をエクセルに記入する")
        ],
        [
            sg.Button("記入",size=(25,1)),sg.Button("終了",size=(25,1))
        ],
    ] , size=(520, 630)
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
        selected_file()
        fig = make_data_fig(fig, make=False)
        fig = make_data_fig_selected(fig)
        fig_agg.draw()


    elif event == '-display-':
        fig = make_data_fig(fig, make=True)
        fig_agg.draw()

    elif event == '-clear-':
        fig = make_data_fig(fig, make=False)
        fig_agg.draw()

    
    elif event == '-d1-':
        window["-f1-"].Update("ファイルを選択してください")
        if values["-filename-"] != "":
            window["-f1-"].Update("処理を実行しています")
            fig = make_data_fig(fig, make=False)
            fig = make_data_fig(fig, make=True)
            fig_agg.draw()
        if values['-d1-'] == False:
            window["-f1-"].Update(txt1) 
    elif event == '-d2-':
        window["-f2-"].Update("ファイルを選択してください")
        if values["-filename-"] != "":
            window["-f2-"].Update("処理を実行しています")
            fig = make_data_fig(fig, make=False)
            fig = make_data_fig(fig, make=True)
            fig_agg.draw()
        if values['-d2-'] == False:
            window["-f2-"].Update(txt1) 
    elif event == '-d3-':
        window["-f3-"].Update("ファイルを選択してください")
        if values["-filename-"] != "":
            window["-f3-"].Update("処理を実行しています")
            fig = make_data_fig(fig, make=False)
            fig = make_data_fig(fig, make=True)
            fig_agg.draw()
        if values['-d3-'] == False:
            window["-f3-"].Update(txt1) 
    elif event == '-d4-':
        window["-f4-"].Update("ファイルを選択してください")
        if values["-filename-"] != "":
            window["-f4-"].Update("処理を実行しています")
            fig = make_data_fig(fig, make=False)
            fig = make_data_fig(fig, make=True)
            fig_agg.draw()
        if values['-d4-'] == False:
            window["-f4-"].Update(txt1) 
    elif event == '-d5-':
        window["-f5-"].Update("ファイルを選択してください")
        if values["-filename-"] != "":
            window["-f5-"].Update("処理を実行しています")
            fig = make_data_fig(fig, make=False)
            fig = make_data_fig(fig, make=True)
            fig_agg.draw()
        if values['-d5-'] == False:
            window["-f5-"].Update(txt1) 
    elif event == '-d6-':
        window["-f6-"].Update("ファイルを選択してください")
        if values["-filename-"] != "":
            window["-f6-"].Update("処理を実行しています")
            fig = make_data_fig(fig, make=False)
            fig = make_data_fig(fig, make=True)
            fig_agg.draw() 
        if values['-d6-'] == False: 
            window["-f6-"].Update(txt1)
    elif event == '-d7-':
        window["-f7-"].Update("ファイルを選択してください")
        if values["-filename-"] != "":
            window["-f7-"].Update("処理を実行しています")
            fig = make_data_fig(fig, make=False)
            fig = make_data_fig(fig, make=True)
            fig_agg.draw() 
        if values['-d7-'] == False: 
            window["-f7-"].Update(txt1)
    elif event == '-d8-':
        window["-f8-"].Update("ファイルを選択してください")
        if values["-filename-"] != "":
            window["-f8-"].Update("処理を実行しています")
            fig = make_data_fig(fig, make=False)
            fig = make_data_fig(fig, make=True)
            fig_agg.draw() 
        if values['-d8-'] == False: 
            window["-f8-"].Update(txt1)
    elif event == '-d9-':
        window["-f9-"].Update("ファイルを選択してください")
        if values["-filename-"] != "":
            window["-f9-"].Update("処理を実行しています")
            fig = make_data_fig(fig, make=False)
            fig = make_data_fig(fig, make=True)
            fig_agg.draw() 
        if values['-d9-'] == False: 
            window["-f9-"].Update(txt1)            

window.close()