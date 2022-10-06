from fileinput import filename
from msilib.schema import CheckBox
from tabnanny import filename_only
import PySimpleGUI as sg
import numpy as np 
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import pandas as pd
import pyperclip


ds=[]
for i in range(1,10):
    ds.append(["-d"+str(i)+"-",i,"D="+str(i),"-f"+str(i)+"-","-cp"+str(i)+"-"])
d_values = []
for d in ds:
    d_values.append(d[0])


x_latent = range(0,101)
txt1="ここに数式が出ます" 
txt2="ファイルを選択してください"
txt3="処理を実行しています"

fig = plt.figure(figsize=(5, 4))
ax = fig.add_subplot(111)


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
        for i in range(1,10):
            polyfit_lst.append(list(np.polyfit(xs,ys,i)))

        xl_poly_lst=[]
        output_poly_lst=[]
        for polyfit in polyfit_lst:
            a = len(polyfit)-1
            xl_poly = ""
            output_poly = ""
            for c in polyfit:
                m = 0
                if c < 0:
                    m = 1
                    if xl_poly != "":
                        xl_poly = xl_poly[:-1]
                        output_poly = output_poly[:-1]
                xl_poly += str(c)
                cs = []
                r = 2
                if "e" in str(c):
                    k = str(c).find("e")
                    print("k=",k)
                    
                    output_poly += str(c)[:4+m] + str(c)[k:]
                else:
                    for n in str(c):
                        if n !="." and n != "-" :
                            cs.append(n)
                    for n in cs:
                        if n == '0':
                            r += 1
                            print(r)
                        else:
                            break
                    print(cs)
                    output_poly += str(round(c,r))
                    print("r=",r)
                    r = 2
                if a == 0:
                    xl_poly_lst.append(xl_poly)
                    output_poly_lst.append(output_poly)
                    print(xl_poly)
                    print(output_poly)
                elif a == 1:
                    xl_poly += "x+"
                    output_poly += "x+"
                else:
                    xl_poly += "x^"+str(a)+"+"
                    output_poly += "x^"+str(a)+"+"
                a -= 1

        ax.scatter(xs, ys,label="observed")
        x_latent = np.linspace(min(xs), max(xs), 100)


        for d in ds:
            if values[d[0]]:
                fitted_curve = np.poly1d((lambda x, y: np.polyfit(x, y, d[1]))(xs, ys))(x_latent)
                ax.plot(x_latent, fitted_curve, label=d[2])
                window[d[3]].Update(output_poly_lst[d[1]-1])



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

def selected_checkbox(checkbox_v):
    window[ds[checkbox_v][3]].Update(txt2)
    if values["-filename-"] != "":
        window[ds[checkbox_v][3]].Update(txt3)
        fig = plt.figure(figsize=(5, 4))
        fig = make_data_fig(fig, make=False)
        fig = make_data_fig(fig, make=True)
        fig_agg.draw()
        if values[ds[checkbox_v][0]] == False:
            window[ds[checkbox_v][3]].Update(txt1) 

def copy_clip(i):
    pyperclip.copy(values[ds[i][3]])



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
    col += [[sg.Checkbox(d[2], enable_events=True, key=d[0]),sg.InputText("ここに数式が出ます",readonly=True,key=d[3],size=(80,1)),sg.Button("copy",key=d[4])]]

frame2 = sg.Frame('',
    [    
        [sg.Column(col)],
        [
            sg.Text("結果をエクセルに記入する")
        ],
        [
            sg.Button("記入",size=(25,1)),sg.Button("終了",size=(25,1))
        ],
    ] , size=(1000, 630)
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
        selected_checkbox(0)
    elif event == '-d2-':
        selected_checkbox(1)
    elif event == '-d3-':
        selected_checkbox(2)
    elif event == '-d4-':
        selected_checkbox(3)
    elif event == '-d5-':
        selected_checkbox(4)
    elif event == '-d6-':
        selected_checkbox(5)
    elif event == '-d7-':
        selected_checkbox(6)
    elif event == '-d8-':
        selected_checkbox(7)
    elif event == '-d9-':
        selected_checkbox(8)

    elif event == '-cp1-':
        copy_clip(0)
    elif event == '-cp2-':
        copy_clip(1)
    elif event == '-cp3-':
        copy_clip(2)
    elif event == '-cp4-':
        copy_clip(3)
    elif event == '-cp5-':
        copy_clip(4)
    elif event == '-cp6-':
        copy_clip(5)
    elif event == '-cp7-':
        copy_clip(6)
    elif event == '-cp8-':
        copy_clip(7)
    elif event == '-cp9-':
        copy_clip(8)

         

window.close()