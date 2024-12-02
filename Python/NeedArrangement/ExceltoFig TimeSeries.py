import numpy as np
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

legendinfo = ['Raman Peak', 'Background']
legenditer = iter(legendinfo)
colorstyle = plt.cm.rainbow
alphavalue = 0.7

fd = pathlib.Path(__file__).parent.resolve()
tkfont = {'Font': 'Calibri', 'FontSize': 10}
tickfontstyle = {'Font': 'Calibri', 'FontSize': 18}
fontstyle = {'Font': 'Calibri', 'FontSize': 24}
DefaultInfos = {'Title': 'Title', 'xAxisTitle': 'x-Axis Title', 'yAxisTitle': 'y-Axis Title',
                'xLim': (0, 1), 'yLim': (0, 1), 'tLim': (0, 1), 'Peak':(0, 1), 'MajorTickX': 1, 'MajorTickY': 1}

fh = 450
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

        self.OutputPlotFrame = Frame(self.OutputFrame, bg='white', width=100*fs[0], height=95*fs[1])
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

        TimeLimLable = Label(self.InputInfoFrame, width=14, height=2, text="Time Range", relief="ridge", font=(f"{tkfont['Font']} {tkfont['FontSize']}"))
        TimeLimLable.grid(row=5, column=0, padx=2, pady=2)
        ExpectedPeakLable = Label(self.InputInfoFrame, width=14, height=2, text="Expected Peak, FWHM", relief="ridge", font=(f"{tkfont['Font']} {tkfont['FontSize']}"))
        ExpectedPeakLable.grid(row=6, column=0, padx=2, pady=2)

        MajorTickLable = Label(self.InputInfoFrame, width=14, height=2, text="MajorTick X Y", relief="ridge", font=(f"{tkfont['Font']} {tkfont['FontSize']}"))
        MajorTickLable.grid(row=7, column=0, padx=2, pady=2)
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
        self.xLimEntryDN.grid(row=3, column=1, columnspan=1,padx=2, pady=2)
        self.xLimEntryDN.insert(0, DefaultInfos['xLim'][0])

        self.xLimEntryUP = Entry(self.InputInfoFrame, width=8, textvariable="", relief="ridge", font=(f"{tkfont['Font']} {tkfont['FontSize']}"))
        self.xLimEntryUP.grid(row=3, column=2, columnspan=1, padx=2, pady=2)
        self.xLimEntryUP.insert(0, DefaultInfos['xLim'][1])

        self.yLimEntryDN = Entry(self.InputInfoFrame, width=8, textvariable="", relief="ridge", font=(f"{tkfont['Font']} {tkfont['FontSize']}"))
        self.yLimEntryDN.grid(row=4, column=1, columnspan=1, padx=2, pady=2)
        self.yLimEntryDN.insert(0, DefaultInfos['yLim'][0])

        self.yLimEntryUP = Entry(self.InputInfoFrame, width=8, textvariable="", relief="ridge", font=(f"{tkfont['Font']} {tkfont['FontSize']}"))
        self.yLimEntryUP.grid(row=4, column=2, columnspan=1, padx=2, pady=2)
        self.yLimEntryUP.insert(0, DefaultInfos['yLim'][1])

        self.TimeLimEntryDN = Entry(self.InputInfoFrame, width=8, textvariable="", relief="ridge", font=(f"{tkfont['Font']} {tkfont['FontSize']}"))
        self.TimeLimEntryDN.grid(row=5, column=1, padx=2, pady=2)
        self.TimeLimEntryDN.insert(0, DefaultInfos['tLim'][0])

        self.TimeLimEntryUP = Entry(self.InputInfoFrame, width=8, textvariable="", relief="ridge", font=(f"{tkfont['Font']} {tkfont['FontSize']}"))
        self.TimeLimEntryUP.grid(row=5, column=2, padx=2, pady=2)
        self.TimeLimEntryUP.insert(0, DefaultInfos['tLim'][1])

        self.ExpectedPeakEntry = Entry(self.InputInfoFrame, width=8, textvariable="", relief="ridge", font=(f"{tkfont['Font']} {tkfont['FontSize']}"))
        self.ExpectedPeakEntry.grid(row=6, column=1, padx=2, pady=2)
        self.ExpectedPeakEntry.insert(0, DefaultInfos['Peak'][0])

        self.ExpectedFWHMEntry = Entry(self.InputInfoFrame, width=8, textvariable="", relief="ridge", font=(f"{tkfont['Font']} {tkfont['FontSize']}"))
        self.ExpectedFWHMEntry.grid(row=6, column=2, padx=2, pady=2)
        self.ExpectedFWHMEntry.insert(0, DefaultInfos['Peak'][1])


        self.MajorTickEntryX = Entry(self.InputInfoFrame, width=8, textvariable="", relief="ridge", font=(f"{tkfont['Font']} {tkfont['FontSize']}"))
        self.MajorTickEntryX.grid(row=7, column=1, columnspan=1, padx=2, pady=2)
        self.MajorTickEntryX.insert(0, DefaultInfos['MajorTickX'])

        self.MajorTickEntryY = Entry(self.InputInfoFrame, width=8, textvariable="", relief="ridge", font=(f"{tkfont['Font']} {tkfont['FontSize']}"))
        self.MajorTickEntryY.grid(row=7, column=2, columnspan=1, padx=2, pady=2)
        self.MajorTickEntryY.insert(0, DefaultInfos['MajorTickY'])

        DrawClipboard = Button(self.OutputFrame, width=20, height=2, text="Clipboard to Figure", relief="raised", font=(f"{tkfont['Font']} {tkfont['FontSize']}"), command=self.DrawCurve)
        DrawClipboard.grid(row=1, column=0, rowspan=2, padx=2, pady=2)

        NewFigure = Button(self.OutputFrame, width=20, height=2, text="New Figure", relief="raised", font=(f"{tkfont['Font']} {tkfont['FontSize']}"), command=self.NewFig)
        NewFigure.grid(row=1, column=1, rowspan=2, padx=2, pady=2)

        self.FaceColorXVal = Entry(self.OutputFrame, width=8, textvariable="", relief="ridge", font=(f"{tkfont['Font']} {tkfont['FontSize']}"))
        self.FaceColorXVal.grid(row=1, column=2, padx=1, pady=1)
        self.FaceColorXVal.insert(0, 0)

        PaintFaceColor = Button(self.OutputFrame, width=8, text="Paint", relief="raised", font=(f"{tkfont['Font']} {tkfont['FontSize']}"), command=self.PaintFace)
        PaintFaceColor.grid(row=2, column=2, padx=1, pady=1)

        # SaveFigure = Button(self.OutputFrame, width=14, height=2, text="Save Figure", relief="raised", font=(f"{tkfont['Font']} {tkfont['FontSize']}"), command=self.SaveFigure)
        # SaveFigure.grid(row=1, column=2, padx=2, pady=2)

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
        DefaultInfos['tLim'] = (float(self.TimeLimEntryDN.get()), float(self.TimeLimEntryUP.get()))
        DefaultInfos['Peak'] = (float(self.ExpectedPeakEntry.get()), float(self.ExpectedFWHMEntry.get()))

    def MakeFigure(self):
        self.fig, self.ax = plt.subplots(figsize=(fs[0], fs[1]))
        self.output_plt = FigureCanvasTkAgg(self.fig, self.OutputPlotFrame)
        self.output_plt.get_tk_widget().pack(side=LEFT, fill=BOTH, expand=1)
        plt.close(self.fig)

    def FigureOptionSetting(self, ax):
        ax.set_title(DefaultInfos['Title'], font=fontstyle['Font'], fontsize=fontstyle['FontSize'])
        ax.set_xlabel(DefaultInfos['xAxisTitle'], font=fontstyle['Font'], fontsize=fontstyle['FontSize'])
        ax.set_ylabel(DefaultInfos['yAxisTitle'], font=fontstyle['Font'], fontsize=fontstyle['FontSize'])
        ax.set_xlim(DefaultInfos['tLim'][0], DefaultInfos['tLim'][1])
        ax.set_ylim(DefaultInfos['yLim'][0], DefaultInfos['yLim'][1])
        ax.xaxis.set_major_locator(MultipleLocator(DefaultInfos['MajorTickX']))
        ax.yaxis.set_major_locator(MultipleLocator(DefaultInfos['MajorTickY']))

        ax.grid(True)
        ax.tick_params(axis='x', labelsize=tickfontstyle['FontSize'])
        ax.tick_params(axis='y', labelsize=tickfontstyle['FontSize'])
        plt.tight_layout()
        self.forceAspect(ax)

    def DrawCurve(self):

        c = next(self.color)
        # l = next(legenditer)

        '1. Data Import and Crop'
        data = self.ReadClipboard()
        x, t, data = self.Apply_xRange(data, DefaultInfos['xLim'])
        t, data = self.Apply_tRange(t, data, DefaultInfos['tLim'])

        '2. Time Average'
        mean, stddev = self.Average(data)

        '3. Find Raman Peak Wavenumber'
        x_Pick, x_max = self.PickIdx(x, mean, stddev, DefaultInfos['Peak'])
        Range = np.arange(-10, 11, 1)
        Peak_Range = x_max + Range
        Flat_Range_Left = x_Pick[0] + Range
        Flat_Range_Right = x_Pick[-1] + Range

        '4. Background Correction'
        peak = x[x_max]
        y_data = np.max(data[Peak_Range[0]:Peak_Range[-1]+1], axis=0)[:]\
                 - (np.average(data[Flat_Range_Left[0]:Flat_Range_Left[-1]+1], axis=0)[:] +
                    np.average(data[Flat_Range_Right[0]:Flat_Range_Right[-1] + 1], axis=0)[:])/2
        x_data = np.ones(y_data.__len__())

        '5. Time Average'
        mu, sigma = np.mean(y_data), np.std(y_data)

        '6. Select data <= mean'
        y_lowdata, t_lowdata = y_data[np.argwhere(y_data<=mu).flatten()], t[np.argwhere(y_data<=mu).flatten()]

        '7. Find Sigma'
        y_low_sigma = np.std(y_lowdata)
        effective_sigma = 2*y_low_sigma

        '8-1. Remove Outlier'
        y_effective_data = y_data[np.argwhere((y_data-mu<=effective_sigma)).flatten()]
        t_effective_data = t[np.argwhere((y_data-mu<=effective_sigma)).flatten()]
        x_effective_data = 2*np.ones(y_effective_data.__len__())
        mu_effective, sigma_effecitve = np.mean(y_effective_data), np.std(y_effective_data)

        '8-2. Only Outlier'
        y_outlier_data = y_data[np.argwhere((y_data-mu>effective_sigma)).flatten()]
        t_outlier_data = t[np.argwhere((y_data-mu>effective_sigma)).flatten()]
        x_outlier_data = 3 * np.ones(y_outlier_data.__len__())
        mu_outlier, sigma_outlier = np.mean(y_outlier_data), np.std(y_outlier_data)


        params = f"$\mathit{{\mu}}={np.round(mu, 1)}, \mathit{{\sigma}}={np.round(sigma, 1)}$"
        # self.drawax.plot(x, mean, c=c, alpha=alphavalue,  marker='o', linestyle='None')
        self.drawax.plot(t, y_data, c=c, alpha=alphavalue, label = params)


        # self.drawax.plot(data[0], data[1], c=c, alpha=alphavalue, label=l, marker='o')
        self.drawax.legend(loc='best', fontsize=fontstyle['FontSize'])
        plt.pause(0.001)

    def NewFig(self):
        self.color = iter(colorstyle(np.linspace(1, 0, legendinfo.__len__())))
        legenditer = iter(legendinfo)
        fig, self.drawax = plt.subplots(figsize=(fs[0], fs[1]))
        self.FigureOptionSetting(self.drawax)
        self.forceAspect(self.drawax)

    def ReadClipboard(self):

        return pd.read_clipboard()

    def forceAspect(self, ax, aspect=1):
        xlim = ax.get_xlim()
        ylim = ax.get_ylim()
        ax.set_aspect(abs((xlim[1] - xlim[0]) / (ylim[1] - ylim[0])) / aspect)

    def PaintFace(self):
        x = float(self.FaceColorXVal.get())
        self.drawax.axvspan(x-3, x+3, facecolor='g', alpha=0.2)

    def Average(self, data):

        Mean, Std = np.zeros(np.shape(data)[0]), np.zeros(np.shape(data)[0])

        for k, data_now in enumerate(data):
            Mean[k] = np.mean(data_now)
            Std[k] = np.std(data_now)

        return Mean, Std

    def PickIdx(self, x, mean, stddev, Peak):
        ExpectedPeak, FWHM = Peak[0], Peak[1]
        xlow = np.argmin(abs(x-(ExpectedPeak-FWHM)))
        xhigh = np.argmin(abs(x-(ExpectedPeak+FWHM)))

        xPeak = xlow + np.argmax(mean[xlow:xhigh])

        idxmin_left = mean[:xlow].argmin()
        idxmin_right = xhigh + mean[xhigh:].argmin()

        idc = [idxmin_left, xPeak, idxmin_right]

        return idc, xPeak

    def Apply_xRange(self, data_raw, xRange):
        x = np.array(data_raw.axes[0], float)
        t = np.array(data_raw.axes[1], float)
        data = np.array(data_raw)
        idxmin = np.argmax(x>xRange[0])
        idxmax = np.argmax(x>xRange[-1])
        return x[idxmin-1:idxmax], t, data[idxmin-1:idxmax]

    def Apply_tRange(self, t, data, tRange):
        t, data = t[np.argwhere(t>=tRange[0]).flatten()], data[:, np.argwhere(t>=tRange[0]).flatten()]
        t, data = t[np.argwhere(t<tRange[-1]).flatten()], data[:, np.argwhere(t<tRange[-1]).flatten()]

        return t, data

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