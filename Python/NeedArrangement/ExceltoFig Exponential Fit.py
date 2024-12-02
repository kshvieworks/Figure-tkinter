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

legendinfo = [r"Raw Data"]
legenditer = iter(legendinfo)
colorstyle = plt.cm.rainbow
alphavalue = 0.7

fd = pathlib.Path(__file__).parent.resolve()
tkfont = {'Font': 'Calibri', 'FontSize': 10}
tickfontstyle = {'Font': 'Calibri', 'FontSize': 18}
fontstyle = {'Font': 'Calibri', 'FontSize': 24}
DefaultInfos = {'Title': 'Title', 'xAxisTitle': 'x-Axis Title', 'yAxisTitle': 'y-Axis Title',
                'xLim': (0, 1), 'yLim': (0, 1), 'MajorTickX': 1, 'MajorTickY': 1}

# tau = 0.05
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

        DrawClipboard = Button(self.OutputFrame, width=18, height=2, text="Clipboard to Figure", relief="raised", font=(f"{tkfont['Font']} {tkfont['FontSize']}"), command=self.DrawCurve)
        DrawClipboard.grid(row=1, column=0, padx=2, pady=2)

        NewFigure = Button(self.OutputFrame, width=18, height=2, text="New Figure", relief="raised", font=(f"{tkfont['Font']} {tkfont['FontSize']}"), command=self.NewFig)
        NewFigure.grid(row=1, column=1, padx=2, pady=2)

        # ApplyLPFBTN = Button(self.OutputFrame, width=7, height=2, text="Apply LPF", relief="raised",
        #                          font=(f"{tkfont['Font']} {tkfont['FontSize']}"), command=self.ApplyLPF)
        # ApplyLPFBTN.grid(row=1, column=2)

        DrawExponential = Button(self.OutputFrame, width=14, height=2, text="Exponential Fitting", relief="raised", font=(f"{tkfont['Font']} {tkfont['FontSize']}"), command=self.DrawExponentialCurve)
        DrawExponential.grid(row=1, column=2)

    def Applyinfo(self):

        self.UpdateInfos()

        if not hasattr(self, 'ax'):
            self.MakeFigure()
            self.NewFig()
            self.color = iter(colorstyle(np.linspace(1, 0, legendinfo.__len__())))
            self.color2 = iter(colorstyle(np.linspace(0, 1, legendinfo.__len__())))
            self.legenditer = iter(legendinfo)

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
        l = next(self.legenditer)

        # self.drawax.plot(self.data[0], self.data[1], c=c, alpha=alphavalue, label=l)
        self.drawax.plot(self.data[0], self.data[1], c=c, alpha=alphavalue/4, marker='o', linestyle = 'None', markeredgecolor='None', label = l)
        self.drawax.legend(loc='best', fontsize=tkfont['FontSize'])
        plt.pause(0.001)

    def NewFig(self):
        self.color = iter(colorstyle(np.linspace(1, 0, legendinfo.__len__())))
        self.color2 = iter(colorstyle(np.linspace(0, 1, legendinfo.__len__())))
        self.legenditer = iter(legendinfo)
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

    def DrawExponentialCurve(self):

        self.calData = self.data.copy()
        self.calData[0] = self.calData[0] - self.calData[0][0]

        lc = next(self.color2)

        popt, _ = curve_fit(self.Exponential_Curve, self.calData[0], self.calData[1],
                            [(self.calData[1][0]-self.calData.iloc[-1][1]), -10, self.calData.iloc[-1][1]], maxfev = 10000)
        R2 = self.RSquared(self.calData[0], self.calData[1], self.Exponential_Curve(self.calData[0], *popt))

        x = np.linspace(np.min(self.calData[0]), np.max(self.calData[0]), 2000)
        fit_data = self.Exponential_Curve(x, *popt)

        params = f"f(x) = {popt[0]:.2e}$\\times$$e^\\frac{{t-{np.round(self.data[0][0], 1)}}}{{{np.round(1/popt[1], 2)}}}$ + {popt[2]:.2e} " + \
                 f"\n $t_r$ = {np.abs(np.round(1/popt[1]*np.log(9), 2))} s" + \
                 f"\n R$^2$ = {np.round(R2, 3)}"

        self.drawax.plot(x+self.data[0][0], fit_data, c=lc, alpha=alphavalue, label=params)

        self.drawax.legend(loc='best', fontsize=tickfontstyle['FontSize'] - 4)
        plt.pause(0.001)

        asdf = 1

    # def ApplyLPF(self):
    #     self.data[1] = self.LPF_1stOrder(self.data[1], tau, self.data[0][1]-self.data[0][0])
    #     self.drawax.plot(self.data[0], self.data[1], 'b--', alpha=alphavalue/4, label = 'LPF')
    #     self.drawax.legend(loc='best', fontsize=tkfont['FontSize'])
    #     plt.pause(0.001)
    #
    # def LPF_1stOrder(self, data, tau, dt):
    #     v0 = data.copy()
    #     for k, vi in enumerate(data[1:]):
    #         v0[k + 1] = self._Equation_LPF_1st(vi, v0[k], tau, dt)
    #
    # def _Equation_LPF_1st(self, vi, vo_prev, tau, dt):
    #     return ((dt*vi) + (tau*vo_prev)) / (dt + tau)

    def Exponential_Curve(self, x, A, B, C):
        return A * np.exp(B*x) + C

    def RSquared(self, x, y, fit_data):

        yhat = fit_data
        ybar = np.sum(y) / len(y)
        sse = np.sum((yhat - ybar)**2)
        sst = np.sum((y - ybar)**2)

        return sse/sst


if __name__ == '__main__':
    window = Tk()
    ClipboardtoFig(window)
    window.mainloop()