import numpy as np
import pandas as pd

import xlsxwriter

import matplotlib.pyplot as plt
from matplotlib.ticker import MultipleLocator
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from matplotlib import ticker
import matplotlib.image as mImage

import tkinter
from tkinter import *
import pathlib
from os.path import isfile, join

colorstyle = plt.cm.RdBu_r
alphavalue = 1

fd = pathlib.Path(__file__).parent.resolve()
tkfont = {'Font': 'Calibri', 'FontSize': 10}
tickfontstyle = {'Font': 'Calibri', 'FontSize': 16}
fontstyle = {'Font': 'Calibri', 'FontSize': 24}
DefaultInfos = {'Title': 'Title', 'xAxisTitle': 'x-Axis Title', 'yAxisTitle': 'y-Axis Title',
                'xLim': (0, 1), 'yLim': (0, 1), 'MajorTickX': 1, 'MajorTickY': 1, 'CMapMin': 0, 'CMapMax': 1, 'CMapTitle': 'ColorMap Title'}

fh = 400
fw = fh*2
fs = (fw/200, 0.7*fh/100)

class ClipboardtoFig:
    def __init__(self, window):
        self.window = window
        self.window.title("Clipboard to Figure")
        # self.window.config(background='#FFFFFF')
        self.window.geometry(f"{fw}x{fh}")
        self.window.resizable(False, False)

        self.filepath = ""

        self.__main__()

    def __main__(self):

        self.InputInfoFrame = LabelFrame(self.window, width=fw, height=fh, text="Plot 속성", font=(f"{fontstyle['Font']} {fontstyle['FontSize']}"))
        self.InputInfoFrame.grid(column=0, row=0, padx=10, pady=10)

        self.OutputFrame = LabelFrame(self.window, width=fw, height=fh, text="Figure Property Preview", font=(f"{fontstyle['Font']} {fontstyle['FontSize']}"))
        self.OutputFrame.grid(column=1, row=0, padx=10, pady=10)

        self.OutputPlotFrame = Frame(self.OutputFrame, bg='white', width=100*fs[0], height=100*fs[1])
        self.OutputPlotFrame.grid(column=0, row=0, columnspan=2, padx=10, pady=10)

        TitleLable = Label(self.InputInfoFrame, width=14, height=2, text="Title", relief="ridge", font=(f"{tkfont['Font']} {tkfont['FontSize']}"))
        TitleLable.grid(row=0, column=0, padx=2, pady=2)
        xAxisLable = Label(self.InputInfoFrame, width=14, height=2, text="x-Axis Title ", relief="ridge", font=(f"{tkfont['Font']} {tkfont['FontSize']}"))
        xAxisLable.grid(row=1, column=0, padx=2, pady=2)
        yAxisLable = Label(self.InputInfoFrame, width=14, height=2, text="y-Axis Title", relief="ridge", font=(f"{tkfont['Font']} {tkfont['FontSize']}"))
        yAxisLable.grid(row=2, column=0, padx=2, pady=2)
        xLimLable = Label(self.InputInfoFrame, width=14, height=2, text="xLim", relief="ridge", font=(f"{tkfont['Font']} {tkfont['FontSize']}"))
        xLimLable.grid(row=3, column=0, padx=2, pady=2)
        yLimLable = Label(self.InputInfoFrame, width=14, height=2, text="yLim", relief="ridge", font=(f"{tkfont['Font']} {tkfont['FontSize']}"))
        yLimLable.grid(row=4, column=0, padx=2, pady=2)
        MajorTickLable = Label(self.InputInfoFrame, width=14, height=2, text="MajorTick X Y", relief="ridge", font=(f"{tkfont['Font']} {tkfont['FontSize']}"))
        MajorTickLable.grid(row=5, column=0, padx=2, pady=2)
        CMapLimLable = Label(self.InputInfoFrame, width=14, height=1, text="CMap Lim", relief="ridge", font=(f"{tkfont['Font']} {tkfont['FontSize']}"))
        CMapLimLable.grid(row=6, column=0, padx=2, pady=2)
        CMapLable = Label(self.InputInfoFrame, width=14, height=1, text="CMap Title", relief="ridge", font=(f"{tkfont['Font']} {tkfont['FontSize']}"))
        CMapLable.grid(row=7, column=0, padx=2, pady=2)
        ApplyInfo = Button(self.InputInfoFrame, width=14, height=2, text="Apply", relief="raised", font=(f"{tkfont['Font']} {tkfont['FontSize']}"), command=self.Applyinfo)
        ApplyInfo.grid(row=8, column=0, columnspan=3, padx=2, pady=2)

        self.TitleEntry = Entry(self.InputInfoFrame, width=20, relief="ridge", font=(f"{tkfont['Font']} {tkfont['FontSize']}"))
        self.TitleEntry.grid(row=0, column=1, columnspan=2, padx=2, pady=2)
        self.TitleEntry.insert(0, DefaultInfos['Title'])

        self.xAxisEntry = Entry(self.InputInfoFrame, width=20, textvariable="", relief="ridge", font=(f"{tkfont['Font']} {tkfont['FontSize']}"))
        self.xAxisEntry.grid(row=1, column=1, columnspan=2, padx=2, pady=2)
        self.xAxisEntry.insert(0, DefaultInfos['xAxisTitle'])

        self.yAxisEntry = Entry(self.InputInfoFrame, width=20, textvariable="", relief="ridge", font=(f"{tkfont['Font']} {tkfont['FontSize']}"))
        self.yAxisEntry.grid(row=2, column=1, columnspan=2, padx=2, pady=2)
        self.yAxisEntry.insert(0, DefaultInfos['yAxisTitle'])

        self.xLimEntryDN = Entry(self.InputInfoFrame, width=8, textvariable="", relief="ridge", font=(f"{tkfont['Font']} {tkfont['FontSize']}"))
        self.xLimEntryDN.grid(row=3, column=1, padx=2, pady=2)
        self.xLimEntryDN.insert(0, DefaultInfos['xLim'][0])

        self.xLimEntryUP = Entry(self.InputInfoFrame, width=8, textvariable="", relief="ridge", font=(f"{tkfont['Font']} {tkfont['FontSize']}"))
        self.xLimEntryUP.grid(row=3, column=2, padx=2, pady=2)
        self.xLimEntryUP.insert(0, DefaultInfos['xLim'][1])

        self.yLimEntryDN = Entry(self.InputInfoFrame, width=8, textvariable="", relief="ridge", font=(f"{tkfont['Font']} {tkfont['FontSize']}"))
        self.yLimEntryDN.grid(row=4, column=1, padx=2, pady=2)
        self.yLimEntryDN.insert(0, DefaultInfos['yLim'][0])

        self.yLimEntryUP = Entry(self.InputInfoFrame, width=8, textvariable="", relief="ridge", font=(f"{tkfont['Font']} {tkfont['FontSize']}"))
        self.yLimEntryUP.grid(row=4, column=2, padx=2, pady=2)
        self.yLimEntryUP.insert(0, DefaultInfos['yLim'][1])

        self.MajorTickEntryX = Entry(self.InputInfoFrame, width=8, textvariable="", relief="ridge", font=(f"{tkfont['Font']} {tkfont['FontSize']}"))
        self.MajorTickEntryX.grid(row=5, column=1, padx=2, pady=2)
        self.MajorTickEntryX.insert(0, DefaultInfos['MajorTickX'])

        self.MajorTickEntryY = Entry(self.InputInfoFrame, width=8, textvariable="", relief="ridge", font=(f"{tkfont['Font']} {tkfont['FontSize']}"))
        self.MajorTickEntryY.grid(row=5, column=2, padx=2, pady=2)
        self.MajorTickEntryY.insert(0, DefaultInfos['MajorTickY'])

        self.CMapEntryMin = Entry(self.InputInfoFrame, width=8, textvariable="", relief="ridge", font=(f"{tkfont['Font']} {tkfont['FontSize']}"))
        self.CMapEntryMin.grid(row=6, column=1, padx=2, pady=2)
        self.CMapEntryMin.insert(0, DefaultInfos['CMapMin'])

        self.CMapEntryMax = Entry(self.InputInfoFrame, width=8, textvariable="", relief="ridge", font=(f"{tkfont['Font']} {tkfont['FontSize']}"))
        self.CMapEntryMax.grid(row=6, column=2, padx=2, pady=0)
        self.CMapEntryMax.insert(0, DefaultInfos['CMapMax'])

        self.CMapEntry = Entry(self.InputInfoFrame, width=20, textvariable="", relief="ridge", font=(f"{tkfont['Font']} {tkfont['FontSize']}"))
        self.CMapEntry.grid(row=7, column=1, columnspan=2, padx=2, pady=0)
        self.CMapEntry.insert(0, DefaultInfos['CMapTitle'])

        DrawClipboard = Button(self.OutputFrame, width=38, height=2, text="Clipboard to Figure", relief="raised", font=(f"{tkfont['Font']} {tkfont['FontSize']}"), command=self.DrawFig)
        DrawClipboard.grid(row=1, column=0, padx=2, pady=2)
        SaveFigure = Button(self.OutputFrame, width=14, height=2, text="Save Figure", relief="raised", font=(f"{tkfont['Font']} {tkfont['FontSize']}"), command=self.SaveFigure)
        SaveFigure.grid(row=1, column=1, columnspan=2, padx=2, pady=2)

    def Applyinfo(self):

        self.UpdateInfos()

        if not hasattr(self, 'ax'):
            self.MakeFigure()

        self.ax.cla()
        self.FigureOptionSetting(self.ax)

    def UpdateInfos(self):
        DefaultInfos['Title'] = self.TitleEntry.get()
        DefaultInfos['xAxisTitle'] = self.xAxisEntry.get()
        DefaultInfos['yAxisTitle'] = self.yAxisEntry.get()
        DefaultInfos['xLim'] = (float(self.xLimEntryDN.get()), float(self.xLimEntryUP.get()))
        DefaultInfos['yLim'] = (float(self.yLimEntryDN.get()), float(self.yLimEntryUP.get()))
        DefaultInfos['MajorTickX'] = float(self.MajorTickEntryX.get())
        DefaultInfos['MajorTickY'] = float(self.MajorTickEntryY.get())
        DefaultInfos['CMapMin'] = float(self.CMapEntryMin.get())
        DefaultInfos['CMapMax'] = float(self.CMapEntryMax.get())
        DefaultInfos['CMapTitle'] = self.CMapEntry.get()

    def MakeFigure(self):
        self.fig, self.ax = plt.subplots(figsize=(fs[0], fs[1]))
        self.output_plt = FigureCanvasTkAgg(self.fig, self.OutputPlotFrame)
        self.output_plt.get_tk_widget().pack(side=LEFT, fill=BOTH, expand=1)
        # plt.close(fig)

    def FigureOptionSetting(self, ax):
        ax.set_title(DefaultInfos['Title'], font=fontstyle['Font'], fontsize=fontstyle['FontSize'])
        ax.set_xlabel(DefaultInfos['xAxisTitle'], font=fontstyle['Font'], fontsize=fontstyle['FontSize'])
        ax.set_ylabel(DefaultInfos['yAxisTitle'], font=fontstyle['Font'], fontsize=fontstyle['FontSize'])
        ax.set_xlim(DefaultInfos['xLim'][0], DefaultInfos['xLim'][1])
        ax.set_ylim(DefaultInfos['yLim'][0], DefaultInfos['yLim'][1])
        ax.xaxis.set_major_locator(MultipleLocator(DefaultInfos['MajorTickX']))
        ax.yaxis.set_major_locator(MultipleLocator(DefaultInfos['MajorTickY']))

        ax.grid(True)
        ax.tick_params(axis='x', labelsize=tickfontstyle['FontSize'])
        ax.tick_params(axis='y', labelsize=tickfontstyle['FontSize'])
        plt.tight_layout()
        self.forceAspect(ax)

    def DrawFig(self):

        fig, ax = plt.subplots(figsize=(fs[0], fs[1]))
        self.FigureOptionSetting(ax)

        data = self.ReadClipboard(0)
        yrange = data.index.values
        xrange = pd.to_numeric(data.columns.values, errors='coerce')
        c = ax.imshow(data, cmap=colorstyle, alpha=alphavalue, extent=[xrange[0], xrange[-1], yrange[0], yrange[-1]],
                      origin='lower', vmin=DefaultInfos['CMapMin'], vmax=DefaultInfos['CMapMax'])
        ax.cbar = fig.colorbar(c, ax=ax)
        ax.cbar.set_label(label=DefaultInfos['CMapTitle'], size=fontstyle['FontSize'])
        ax.cbar.ax.tick_params(labelsize=tickfontstyle['FontSize'])
        ax.cbar.locator = ticker.MaxNLocator(nbins=2)
        ax.cbar.update_ticks()

        self.forceAspect(ax)
        plt.pause(0.001)
        plt.show()

    def ReadClipboard(self, headervalue=None):

        return pd.read_clipboard(header=headervalue)

    def forceAspect(self, ax, aspect=1):
        xlim = ax.get_xlim()
        ylim = ax.get_ylim()
        ax.set_aspect(abs((xlim[1] - xlim[0]) / (ylim[1] - ylim[0])) / aspect)

    def SaveFigure(self):

        filepath = tkinter.filedialog.asksaveasfilename(initialdir=f"{fd}/",
                                                        title="Save as",
                                                        filetypes=(("png", ".png"),
                                                                   ("all files", "*")))
        filepath = f"{filepath}.png"

        self.fig.savefig(filepath)

if __name__ == '__main__':
    window = Tk()
    ClipboardtoFig(window)
    window.mainloop()