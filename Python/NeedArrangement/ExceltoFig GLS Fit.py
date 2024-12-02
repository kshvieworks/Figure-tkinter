import numpy as np
from scipy.optimize import curve_fit
import pandas as pd

import xlsxwriter

import matplotlib.pyplot as plt
from matplotlib.ticker import MultipleLocator
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import matplotlib.image as mImage

import tkinter
from tkinter import *
import pathlib
from os.path import isfile, join

legendinfo = ['Measurement 1', 'Measurement 2']
legenditer = iter(legendinfo)
colorstyle = plt.cm.rainbow
alphavalue = 0.4

fd = pathlib.Path(__file__).parent.resolve()
tkfont = {'Font': 'Calibri', 'FontSize': 10}
tickfontstyle = {'Font': 'Calibri', 'FontSize': 18}
fontstyle = {'Font': 'Calibri', 'FontSize': 24}
DefaultInfos = {'Title': 'Title', 'xAxisTitle': 'x-Axis Title', 'yAxisTitle': 'y-Axis Title',
                'xLim': (0, 1), 'yLim': (0, 1), 'MajorTickX': 1, 'MajorTickY': 1}

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
        self.OutputPlotFrame.grid(column=0, row=0, columnspan=3, padx=10, pady=10)

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
        ApplyInfo = Button(self.InputInfoFrame, width=14, height=2, text="Apply", relief="raised", font=(f"{tkfont['Font']} {tkfont['FontSize']}"), command=self.Applyinfo)
        ApplyInfo.grid(row=6, column=0, columnspan=3, padx=2, pady=2)

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

        DrawClipboard = Button(self.OutputFrame, width=20, height=2, text="Clipboard to Figure", relief="raised", font=(f"{tkfont['Font']} {tkfont['FontSize']}"), command=self.DrawCurve)
        DrawClipboard.grid(row=1, column=0, padx=2, pady=2)

        NewFigure = Button(self.OutputFrame, width=20, height=2, text="New Figure", relief="raised", font=(f"{tkfont['Font']} {tkfont['FontSize']}"), command=self.NewFig)
        NewFigure.grid(row=1, column=1, padx=2, pady=2)

        DrawGaussian = Button(self.OutputFrame, width=14, height=2, text="Gaussian Fitting", relief="raised", font=(f"{tkfont['Font']} {tkfont['FontSize']}"), command=self.DrawGaussianCurve)
        DrawGaussian.grid(row=1, column=2, padx=2, pady=2)

    def Applyinfo(self):

        self.UpdateInfos()

        if not hasattr(self, 'ax'):
            self.MakeFigure()
            self.NewFig()
            self.color = iter(colorstyle(np.linspace(1, 0, legendinfo.__len__())))

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

    def MakeFigure(self):
        self.fig, self.ax = plt.subplots(figsize=(fs[0], fs[1]))
        self.output_plt = FigureCanvasTkAgg(self.fig, self.OutputPlotFrame)
        self.output_plt.get_tk_widget().pack(side=LEFT, fill=BOTH, expand=1)
        plt.close(self.fig)

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

    def DrawCurve(self):

        self.data = self.ReadClipboard()
        c = next(self.color)
        l = next(legenditer)

        self.drawax.plot(self.data[0], self.data[1], c=c, alpha=alphavalue, label = l, marker='o')
        # self.drawax.plot(self.data[0], self.data[1], c=c, alpha=alphavalue, label=l, marker='o')
        self.drawax.legend(loc='best', fontsize=tickfontstyle['FontSize'])
        plt.pause(0.001)

    def NewFig(self):
        self.color = iter(colorstyle(np.linspace(1, 0, legendinfo.__len__())))
        legenditer = iter(legendinfo)
        plt.pause(0.001)
        fig, self.drawax = plt.subplots(figsize=(fs[0], fs[1]))
        self.FigureOptionSetting(self.drawax)
        self.forceAspect(self.drawax)

    def ReadClipboard(self):

        return pd.read_clipboard(header=None)

    def forceAspect(self, ax, aspect=1):
        xlim = ax.get_xlim()
        ylim = ax.get_ylim()
        ax.set_aspect(abs((xlim[1] - xlim[0]) / (ylim[1] - ylim[0])) / aspect)

    def Gaussian_Curve(self, x, amplitude, mean, stddev):
        m = 0.95
        return amplitude * np.exp(-(((x - mean) / (np.sqrt(2)*stddev))**2)*(1-m)) \
               * 1 / (1+(((x - mean)/(np.sqrt(2*np.log(2))*stddev))**2)*m)

    def DrawGaussianCurve(self):

        calibrated_y = self.data[1] - np.min(self.data[1])
        Initial_Amplitude, Initial_Mean, Initial_Stddev = max(calibrated_y), self.data[0][np.argmax(self.data[1])], 5
        popt, _ = curve_fit(self.Gaussian_Curve, self.data[0], calibrated_y, [Initial_Amplitude, Initial_Mean, Initial_Stddev])

        x = np.linspace(np.min(self.data[0]), np.max(self.data[0]), 1000)
        fit_data = self.Gaussian_Curve(x, *popt) + np.min(self.data[1])
        # params = f"$\mathit{{\mu}}={np.round(popt[1], 2)}, \mathit{{\sigma}}={np.round(popt[2], 2)}$"
        params = f"$\mathit{{\mu}}={np.round(popt[1], 1)}, \mathit{{\sigma}}={np.rint(popt[2])}$"

        # self.Gaussian_FWHM(fit_data, x)

        self.drawax.plot(x, fit_data, 'g', alpha=0.8, label=params)
        self.drawax.legend(loc='best', fontsize=tkfont['FontSize'])
        plt.pause(0.001)

        asdf = 1

    def Gaussian_FWHM(self, fit_data, x):
        FWHM_Intensity = (np.max(fit_data) - np.min(fit_data)) / 2 + np.min(fit_data)
        d = np.array(np.where(np.sign(FWHM_Intensity - fit_data) < 0)).ravel()
        idx1, idx2, idx3 = d[0], d[-1], np.argmax(fit_data)
        self.drawax.axvspan(x[idx1], x[idx2], facecolor='g', alpha=0.2)
        self.drawax.text(x[idx2], 2*fit_data[idx3]/3, f'FWHM={int(x[idx2]-x[idx1])}', fontsize=fontstyle['FontSize'])


if __name__ == '__main__':
    window = Tk()
    ClipboardtoFig(window)
    window.mainloop()