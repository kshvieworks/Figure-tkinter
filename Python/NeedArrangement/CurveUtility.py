import numpy as np
from scipy.optimize import curve_fit
import pandas as pd

import xlsxwriter

import tkinter.filedialog
import pathlib
from os.path import isfile, join

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
import CurveUtility
import numpy as np
C = CurveUtility.Polynomial_CurveFitting()
data = C.ReadClipboard()
npdata = np.array(data)
xdata = npdata[1:, 0]
ydata = npdata[1:,1:]
x = npdata[0, 1:]

R = np.empty((ydata.shape[0]))
p = np.empty((3, ydata.shape[0], ydata.shape[1]))
for i, idata in enumerate(ydata):

    y = idata
    fit_curve = C.Quadratic_Fitting(x, y)
    R[i] =C.RSquared(x, y, fit_curve(x))
    for k in range(y.__len__()):
        p[:, i, k] = C.PortionCalc(x[k], fit_curve)
ycal = ydata * p[2]
ycal = np.concatenate((ycal, R.reshape(-1, 1)), axis=1)
x = np.append(x, 'R2')

C.Save(xdata, ycal, x)
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

fd = pathlib.Path(__file__).parent.resolve()


class Polynomial_CurveFitting:
    def __init__(self):
        asdf = 1

    def ReadClipboard(self, sep='/s+', index_col=None):
        return pd.read_clipboard(header=None, sep=sep, index_col=index_col)

    def Quadratic_Fitting(self, x, y, i):
        if i > 67:
            asdf = 1

        return np.poly1d(np.polyfit(x, y, 2))

    def RSquared(self, x, y, yhat):
        ybar = np.sum(y) / len(y)
        sse = np.sum((yhat - ybar)**2)
        sst = np.sum((y - ybar)**2)
        return sse/sst

    def PortionCalc(self, x, fit_curve):
        ysum = 0
        for k in range(len(fit_curve.c)):
            ysum += np.abs(fit_curve.coef[k]*np.power(x, fit_curve.order-k))
        y_decomp = np.array([])
        for k in range(len(fit_curve.c)):
            y_decomp = np.append(y_decomp, np.abs(fit_curve.coef[k]*np.power(x, fit_curve.order-k)/ysum))

        return y_decomp

    def Calibration(self, y, CalibrationFactor):
        return y*CalibrationFactor

    def Save(self, x, y, yindex):
        filepath = tkinter.filedialog.asksaveasfilename(initialdir=f"{fd}/",
                                                        title="Save as",
                                                        filetypes=(("csv Files", ".csv"),
                                                                   ("all files", "*")))

        df = pd.DataFrame(index = x, columns = yindex, data = y)
        df.to_csv(f"{filepath}.csv", index=True)

