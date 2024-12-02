import numpy as np
import pandas as pd

import xlsxwriter

import matplotlib as mpl
import matplotlib.pyplot as plt
from matplotlib.ticker import MultipleLocator, Locator
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import matplotlib.image as mImage

import tkinter
from tkinter import *
import pathlib
from os.path import isfile, join

# legendinfo = [f"{chr(0x245F+1)}", f"{chr(0x245F+2)}", f"{chr(0x245F+3)}", f"{chr(0x245F+4)}", f"{chr(0x245F+5)}", f"{chr(0x245F+6)}"]
# legendinfo = ["Design Area", "Practical Area"]
# legendinfo = [f"Sample 3-{chr(0x245F+1)}", f"Sample 2-{chr(0x245F+1)}", f"Sample 1-{chr(0x245F+1)}"]
# legendinfo = [f"Vieworks a-Si PIN PD", f"Dongwoo FineChem OPD 1-{chr(0x245F+1)}", f"Dongwoo FineChem OPD 2-{chr(0x245F+1)}", f"Dongwoo FineChem OPD 3-{chr(0x245F+1)}"]
# legendinfo = [f"QDI Photodiode", f"Vieworks Camera"]
legendinfo = ["Source", "DRZ-Plus", "Sample #1", "Sample #2"]
# legendinfo = ["V$_{EXT}$ = -0.5V", "V$_{EXT}$ = -0.25V"]
# legendinfo = ["V$_{PD}$ = -0.1V"]
# legendinfo = [f"{chr(0x245F+6)}T232T235NN03 (XR3)"]
# legendinfo = ['9-PD18-RL0', '9-PD19-RL0', '9-PD20-RL0', '9-PD27-RL0', '9-PD28-RL0', '9-PD29-RL0']
legenditer = iter(legendinfo)
colorstyle = plt.cm.rainbow
alphavalue = 0.7

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


    def MakeFigure(self):
        self.fig, self.ax = plt.subplots(figsize=(fs[0], fs[1]))
        self.output_plt = FigureCanvasTkAgg(self.fig, self.OutputPlotFrame)
        self.output_plt.get_tk_widget().pack(side=LEFT, fill=BOTH, expand=1)
        self.data1 = np.array([])
        plt.close(self.fig)

    def FigureOptionSetting(self, ax):

        ax.set_title(DefaultInfos['Title'], font=fontstyle['Font'], fontsize=fontstyle['FontSize'])
        ax.set_xlabel(DefaultInfos['xAxisTitle'], font=fontstyle['Font'], fontsize=fontstyle['FontSize'])
        ax.set_ylabel(DefaultInfos['yAxisTitle'], font=fontstyle['Font'], fontsize=fontstyle['FontSize'])
        ax.set_xlim(DefaultInfos['xLim'][0], DefaultInfos['xLim'][1])
        ax.set_ylim(DefaultInfos['yLim'][0], DefaultInfos['yLim'][1])
        ax.xaxis.set_major_locator(MultipleLocator(DefaultInfos['MajorTickX']))

        ax.set_xscale("symlog", linthresh=DefaultInfos['xLim'][0])
        ax.set_yscale("symlog", linthresh=DefaultInfos['yLim'][0])

        # ytickorder = np.arange(np.log10(DefaultInfos['yLim'][1]), np.log10(linth), -np.log10(DefaultInfos['MajorTickY']))
        # ytick = 10 ** ytickorder
        # ytick = np.append(ytick, 0)
        #
        # if DefaultInfos['yLim'][0] < 0:
        #     ytickorder = np.arange(np.log10(abs(DefaultInfos['yLim'][0])), np.log10(linth), -np.log10(DefaultInfos['MajorTickY']))
        #     ytick_Neg = -10 ** ytickorder
        #     ytick = np.append(ytick, ytick_Neg)

        xtickorder = np.arange(np.log10(DefaultInfos['xLim'][1]), np.log10(DefaultInfos['xLim'][0]), -np.log10(DefaultInfos['MajorTickX']))
        if xtickorder[-1] != np.log10(DefaultInfos['xLim'][0]):
            xtickorder = np.append(xtickorder, np.log10(DefaultInfos['xLim'][0]))
        xtick = 10 ** xtickorder
        xtick = np.append(xtick, 0)
        ax.set_xticks(xtick)

        ytickorder = np.arange(np.log10(DefaultInfos['yLim'][1]), np.log10(DefaultInfos['yLim'][0]), -np.log10(DefaultInfos['MajorTickY']))
        if ytickorder[-1] != np.log10(DefaultInfos['yLim'][0]):
            ytickorder = np.append(ytickorder, np.log10(DefaultInfos['yLim'][0]))
        ytick = 10 ** ytickorder
        ytick = np.append(ytick, 0)
        ax.set_yticks(ytick)
        ax.xaxis.set_minor_locator(MinorSymLogLocator(DefaultInfos['xLim'][0], 'xLim'))
        ax.yaxis.set_minor_locator(MinorSymLogLocator(DefaultInfos['yLim'][0], 'yLim'))

        # yminor = LogLocator(base = 10.0, subs = np.arange(1.0, 10.0) * 0.1, numticks = 10)
        # ax.yaxis.set_minor_locator(yminor)

        plt.pause(0.001)

        ax.grid(True)
        ax.tick_params(axis='x', labelsize=tickfontstyle['FontSize'])
        ax.tick_params(axis='y', labelsize=tickfontstyle['FontSize'])
        plt.tight_layout()
        # self.forceAspect(ax)

    def DrawCurve(self):

        data = self.ReadClipboard()
        c = next(self.color)

        # if c[0] == 1:
        self.drawax.plot(data[0], data[1], c=c, alpha=alphavalue, marker='o')
        # else:
        #     self.drawax.plot(data[0], data[1], c=c, alpha=alphavalue, marker='o', linestyle='dotted')
        # self.drawax.plot(data[0], data[1], 'r', alpha=alphavalue)
        # self.drawax.plot(data[0], data[1], c=c, alpha=alphavalue)

        # xval = 600
        # if self.data1.size == 0:
        #     self.yidx = (np.abs(data[0] - xval)).argmin()
        #     self.data1 = np.append(self.data1, data[1][self.yidx])
        # else:
        #
        #     self.TextatBottom(self.drawax, f"J$_{{Design}}$ : J$_{{Practical}}$ = 1 : {np.round(np.abs(data[1][self.yidx])/self.data1[0], 2)}")
        #     self.data1 = np.array([])

        self.drawax.legend(legendinfo, loc='best', fontsize=tickfontstyle['FontSize'])
        plt.pause(0.001)

    def NewFig(self):
        self.color = iter(colorstyle(np.linspace(1, 0, legendinfo.__len__())))
        fig, self.drawax = plt.subplots(figsize=(fs[0], fs[1]))
        self.FigureOptionSetting(self.drawax)
        # self.data1 = np.array([])
        # self.forceAspect(self.drawax)

    def ReadClipboard(self):

        return pd.read_clipboard(header=None)

    def forceAspect(self, ax, aspect=1):
        xlim = ax.get_xlim()
        ylim = np.log10(np.abs(ax.get_ylim()))
        ax.set_aspect(abs((xlim[1] - xlim[0]) / (ylim[1] + ylim[0])) / aspect)

    def PaintFace(self):
        x = float(self.FaceColorXVal.get())
        self.drawax.axvspan(x-10, x+10, facecolor='g', alpha=0.2)

    def TextatTargetPos(self, ax, xval, data):
        yidx = (np.abs(data[0]-xval)).argmin()
        TextHere = f"J$_Design$ : J$_Practical$ = 1 : {np.round(np.abs(data[1][self.yidx])/self.data1[0]), 2}"
        self.drawax.text(xval, data[1][yidx], TextHere, fontsize=fontstyle['FontSize'])

    def TextatBottom(self, ax, TextHere):

        [x1, x2] = ax.get_xlim()
        ax.text((3*x1+x2)/4, DefaultInfos['yLim'][0], TextHere, fontsize=fontstyle['FontSize'])

    def SaveFigure(self):

        filepath = tkinter.filedialog.asksaveasfilename(initialdir=f"{fd}/",
                                                        title="Save as",
                                                        filetypes=(("png", ".png"),
                                                                   ("all files", "*")))
        filepath = f"{filepath}.png"

        self.fig.savefig(filepath)


class MinorSymLogLocator(Locator):
    """
    Dynamically find minor tick positions based on the positions of
    major ticks for a symlog scaling.
    """
    def __init__(self, linthresh, axisinfo):
        """
        Ticks will be placed between the major ticks.
        The placement is linear for x between -linthresh and linthresh,
        otherwise its logarithmically
        """
        self.axisinfo = axisinfo
        self.linthresh = linthresh

    def __call__(self):
        'Return the locations of the ticks'
        # majorlocs = self.axis.get_majorticklocs()
        majorlocs = 10**np.arange(np.log10(DefaultInfos[self.axisinfo][1]), np.log10(DefaultInfos[self.axisinfo][0]) - 1, -1)

        # iterate through minor locs
        minorlocs = []

        # handle the lowest part
        for i in range(1, len(majorlocs)):
            majorstep = majorlocs[i] - majorlocs[i-1]
            if abs(majorlocs[i-1] + majorstep/2) < self.linthresh:
                ndivs = 10
            else:
                ndivs = 9
            minorstep = majorstep / ndivs
            locs = np.arange(majorlocs[i-1], majorlocs[i], minorstep)[1:]
            minorlocs.extend(locs)

        return self.raise_if_exceeds(np.array(minorlocs))

    def tick_values(self, vmin, vmax):
        raise NotImplementedError('Cannot get tick locations for a '
                                  '%s type.' % type(self))


if __name__ == '__main__':
    window = Tk()
    ClipboardtoFig(window)
    window.mainloop()