# -*- coding: utf-8 -*-

import matplotlib
matplotlib.use('TkAgg')

import tkinter as tk
from tkinter.filedialog import askopenfilename, askopenfilenames
import numpy as np
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg, NavigationToolbar2Tk
from collections import OrderedDict
import os.path
from sys import argv
import xlrd as xl
from scipy.stats import linregress
import os, sys
from xlsxwriter import Workbook
from copy import deepcopy


################# WARNINGS #################
# Recent versions of matplotlib and xlrd cause compatibility issues.
# Use with matplotlib 2.x and xlrd 1.x
############################################
RESOLUTION = '960x540'  # Original Resolution of the window
SOFTWARE_NAME = 'Meca Analysis'
CONFIG_FILE_NAME = 'config.dat'
LINE = -1


def locMax(x, y):
    x = abs(x)
    y = abs(y)
    window = 10  # Number of points to average the derivate
    i = 0
    argmax = 0
    deriv = [[], []]
    deriv1, _, _, _, _ = linregress(y[i:i+window], x[i:i+window])
    deriv[0].append(i)
    deriv[1].append(deriv1)
    i += window
    while i < len(x)-window:
        deriv2, _, _, _, _ = linregress(y[i:i+window], x[i:i+window])
        if deriv1 > 0 and deriv2 < 0:
            argmax = np.array(y[i-window:i+window]).argmax() + i-window
            break
        deriv1 = deriv2
        deriv[0].append(i)
        deriv[1].append(deriv1)
        i += window
    return argmax, deriv


def stiffness(deriv, tol):
    angles = [np.arctan(deriv[1][i])*180/np.pi for i in range(len(deriv[1]))]
    bounds = []
    for i in range(len(angles)):
        for j in range(i, len(angles)):
            if abs(angles[j]-angles[i]) > tol:
                break
        bounds.append([i, j])
    argmaxlen = np.array([bounds[i][1]-bounds[i][0] for i in range(len(bounds))]).argmax()
    return [deriv[0][bounds[argmaxlen][0]], deriv[0][bounds[argmaxlen][1]]]



def find_nearest(array_xy, value_xy, aspect_ratio):
    """
    Return the index of the nearest value in array_xy from value_xy.
    Designed to graphically take the nearest value
    """
    x = array_xy[0]
    y = array_xy[1]
    x0 = value_xy[0]
    y0 = value_xy[1]
    a = (max(x) - min(x)) / (max(y) - min(y)) * aspect_ratio
    y = a * y
    y0 = a * y0
    idx = np.sqrt((x - x0)**2 + (y - y0)**2).argmin()
    return idx




class MecaAnalysis(tk.Tk):

    def __init__(self, *args, **kwargs):
        tk.Tk.__init__(self, *args, **kwargs)

        # ATTRIBUTES INITIALIZATION
        self.tol = []
        self.max = []
        self.stiffboundssearch = []
        self.extrapol = []
        self.columnName = []
        self.thisline = None
        self.click = False
        self.controlled_mass_analysis_check = False
        self.loadError = False

        # INITIALIZATION
        self.check_config_file()  # Check if a config file exists. If not create a default one.
        self.import_config_param()  # Import the config parameters

        # DEFINITION OF THE MAIN WINDOW'S NAME AND GEOMETRY
        self.geometry(RESOLUTION)
        self.wm_title(SOFTWARE_NAME)
        info_win = tk.Frame(self)
        info_win.pack(side=tk.LEFT, fill=tk.BOTH)
        right_win = tk.Frame(self)
        right_win.pack(side=tk.RIGHT, fill=tk.BOTH, expand=1)
        canvas_win = tk.Frame(right_win)
        canvas_win.pack(side=tk.TOP, fill=tk.BOTH, expand=1)
        param_win = tk.Frame(right_win)
        param_win.pack(side=tk.BOTTOM)

        # CREATION OF THE PLOT CANVAS
        self.fig = plt.figure(figsize=(0, 0))  # Main figure
        self.ax = self.fig.add_subplot(111)  # Main axis

        self.plotFrame = canvas_win
        self.canvas = FigureCanvasTkAgg(self.fig, master=self.plotFrame)
        self.canvas.draw()
        self.canvas.get_tk_widget().pack(side=tk.TOP, fill=tk.BOTH, expand=1)
        self.toolbar = NavigationToolbar2Tk(self.canvas, self.plotFrame)
        self.toolbar.update()

        # MENU BAR CREATION
        menubar = tk.Menu(self)

        mainMenus = OrderedDict()  # Menu container

            # File menu
        mainMenus['File'] = OrderedDict()
        mainMenus['File']['Load file (Ctrl-l)'] = self.load
        mainMenus['File']['Save file (Ctrl-s)'] = self.save
        mainMenus['File']['Quit'] = self._quit

            # Analysis menu
        mainMenus['Analysis'] = OrderedDict()
        mainMenus['Analysis']['Multiple Files Analysis'] = self.massAnalysis
        mainMenus['Analysis']['Controlled Multiple Files Analysis'] = self.controlledMassAnalyzis

            # Menu bar organization
        for menu in mainMenus:
            curMenu = tk.Menu(menubar, tearoff=0)
            for subMenu in mainMenus[menu]:
                curMenu.add_command(label=subMenu, command=mainMenus[menu][subMenu])
            menubar.add_cascade(label=menu, menu=curMenu)

            # Display the menu
        self.config(menu=menubar)

        # BINDING CTRL-KEY TO THE CORRESPONDING ACTIONS
        self.bind("<Control-l>", self.load)
        self.bind("<Control-s>", self.save)

        # PROPERLY QUIT THE APPLICATION WHEN CLICKING THE CROSS ICON
        self.protocol('WM_DELETE_WINDOW', self._quit)

        # TOOLS AND INFOS PANEL
        self.param_win = param_win
        self.is_data_positive = tk.IntVar()
        self.check_absolute = tk.Checkbutton(self.param_win, text='Force positive data', variable=self.is_data_positive, command=self.force_positive_data)
        self.check_absolute.grid(row=0, column=2, rowspan=1, padx=(10, 3))
        self.linTol_butt = tk.Button(self.param_win, text='Apply', command=self.stiffChange)
        self.linTol_butt.grid(row=1, column=4, rowspan=2, padx=3)
        tk.Label(self.param_win, text='Stiffness tolerance (degree) : ').grid(row=1, column=2, rowspan=2, padx=(10, 3))
        self.linTol_entry = tk.Entry(self.param_win, width=5)
        self.linTol_entry.grid(row=1, column=3, rowspan=2, padx=3)

        self.info_win = info_win
        tk.Label(self.info_win, text='Energy until breakage : ').grid(row=0, column=0, sticky=tk.E, padx=5)
        tk.Label(self.info_win, text='Stiffness : ').grid(row=1, column=0, sticky=tk.E, padx=5)
        tk.Label(self.info_win, text='Real Fmax : ').grid(row=2, column=0, sticky=tk.E, padx=5)
        tk.Label(self.info_win, text='Selected Fmax : ').grid(row=3, column=0, sticky=tk.E, padx=5)
        tk.Label(self.info_win, text='Disp at selected Fmax : ').grid(row=4, column=0, sticky=tk.E, padx=5)
        self.energyval_lab = tk.Label(self.info_win, text='')
        self.energyval_lab.grid(row=0, column=1, sticky=tk.W, padx=5)
        self.stiffnessval_lab = tk.Label(self.info_win, text='')
        self.stiffnessval_lab.grid(row=1, column=1, sticky=tk.W, padx=5)
        self.Fmaxval_lab = tk.Label(self.info_win, text='')
        self.Fmaxval_lab.grid(row=2, column=1, sticky=tk.W, padx=5)
        self.selectedFmax_lab = tk.Label(self.info_win, text='')
        self.selectedFmax_lab.grid(row=3, column=1, sticky=tk.W, padx=5)
        self.selectedDisp_at_Fmax_lab = tk.Label(self.info_win, text='')
        self.selectedDisp_at_Fmax_lab.grid(row=4, column=1, sticky=tk.W, padx=5)

        tk.Button(self.info_win, text='Save', command=self.save).grid(row=5, column=0, columnspan=2, sticky=tk.N+tk.S+tk.E+tk.W)

        if len(argv) > 1:
            self.massAnalysis(filenames=argv[1:])

    def check_config_file(self):
        if not os.path.isfile(CONFIG_FILE_NAME):
            with open('./config.dat', 'w+') as f:
                f.write('tolerance,2,#(in degree) limit of the derivate variation (second derivate). If variation is superior to tolerance, curve is not considered linear anymore.\n')
                f.write('max,1,#if 0 : area between curve is calculated between 0 and the max of the curve. If 1 : area is calculated between 0 and first local max.\n')
                f.write('stiffbounds,0,80,#Boundaries in percent of the max load in which to search the stiffness.\n')
                f.write('extrapol,1,#if 1 : extrapolates the first datas in order to make the curve begin at y=0.\n')
                f.write('columnName,Load,#Name of the column to be taken as the y values\n')
                f.write('startingLine,-1,#Force the number of the line from which the data will be read (if -1, the algorithm does not take it into account).\n')

    def import_config_param(self):
        with open(CONFIG_FILE_NAME, 'r+') as f:
            for line in f:
                line = line.split(',')
                if line[0] == 'tolerance':
                    self.tol = float(line[1])
                if line[0] == 'max':
                    self.max = int(line[1])
                if line[0] == 'stiffbounds':
                    self.stiffboundssearch = [float(line[1]), float(line[2])]
                if line[0] == 'extrapol':
                    self.extrapol = int(line[1])
                if line[0] == 'columnName':
                    self.columnName = line[1]
                if line[0] == 'startingLine':
                    self.startingLine = int(line[1])

    def force_positive_data(self, *args, **kwargs):
        self.load(filename=os.path.join(self.path, self.filename+self.ext))
            

    def load(self, *args, **kwargs):
        """
        Load an excel or csv file and extract the data of load and displacement (specific to Bose testing machines)
        """
        try:
            self.loadError = False
            filepath = kwargs.get('filename', None)
            if filepath is None:
                filepath = askopenfilename(title='Open .xlsx or .csv file from Bose measures',
                                filetypes=[('Excel file, csv file', '*.xlsx *.csv')])
            if not filepath:
                return
            splitpath = os.path.split(filepath)
            self.path = splitpath[0]
            self.filename, self.ext = os.path.splitext(splitpath[1])
            print('\nLoading file : %s'%(self.filename))
            startTime = 0
            if self.ext == '.xlsx':
                rawDatas = xl.open_workbook(filepath).sheet_by_index(0)
                self.Datas = [[], []]
                for i in range(rawDatas.nrows):
                    if rawDatas.row_values(i)[0] == 'Points':
                        if rawDatas.row_values(i-2)[0].lower() == 'rampe retour':
                            continue
                        for j in range(len(rawDatas.row_values(i))):
                            if rawDatas.row_values(i)[j] == 'Elapsed Time':
                                col = j
                                break
                        if float(rawDatas.row_values(i+2)[col]) >= startTime:
                            startTime = float(rawDatas.row_values(i+2)[col])
                            row = i
                block_name =  rawDatas.row_values(row-2)[0]
                print(f"Extracting data from block : {block_name} (starting at line {row})")
                line = rawDatas.row_values(row)
                for i in range(len(line)):
                    if line[i] == 'Disp':
                        colDisp = i
                        self.disp_unit = rawDatas.row_values(row+1)[i]
                    if line[i] == self.columnName:
                        colLoad = i
                        self.col_unit = rawDatas.row_values(row+1)[i]
                print(f"Found displacement column (index {colDisp}): Disp ({self.disp_unit})")
                print(f"Found load column (index {colLoad}): {self.columnName} ({self.col_unit})")
                row += 2
                if self.startingLine != -1:
                    row = self.startingLine - 1
                line = rawDatas.row_values(row)
                while line[col] != '':
                    self.Datas[0].append(float(line[colDisp]))
                    self.Datas[1].append(float(line[colLoad]))
                    row += 1
                    if row >= rawDatas.nrows:
                        break
                    line = rawDatas.row_values(row)
            if self.ext == '.csv':
                self.Datas = [[], []]
                with open(filepath, 'r+') as rawFile:
                    rawDatas = [line.split(',') for line in rawFile]
                    for i in range(len(rawDatas)):
                        if rawDatas[i][0][:len('Points')] == 'Points':
                            for j in range(len(rawDatas[i])):
                                if rawDatas[i][j][:len('Elapsed Time')] == 'Elapsed Time':
                                    col = j
                                    break
                            if float(rawDatas[i+2][col]) >= startTime:
                                startTime = float(rawDatas[i+2][col])
                                row = i
                    line = rawDatas[row]
                    for i in range(len(line)):
                        if line[i][:len('Disp')] == 'Disp':
                            colDisp = i
                            self.disp_unit = rawDatas[row+1][i]
                        if line[i][:len(self.columnName)] == self.columnName:
                            colLoad = i
                            self.col_unit = rawDatas[row+1][i]
                    row += 2
                    if self.startingLine != -1:
                        row = self.startingLine - 1
                    line = rawDatas[row]
                    while line[col] != '':
                        self.Datas[0].append(float(line[colDisp]))
                        self.Datas[1].append(float(line[colLoad]))
                        row += 1
                        if row >= len(rawDatas):
                            break
                        line = rawDatas[row]
            self.Datas[0] = np.array(self.Datas[0]) - self.Datas[0][0]
            self.Datas[1] = np.array(self.Datas[1]) - self.Datas[1][0]
            if self.is_data_positive.get():
                self.Datas[0] = np.abs(self.Datas[0])
                self.Datas[1] = np.abs(self.Datas[1])
            maxArg = abs(self.Datas[1]).argmax()+1

            ruptureArg = -1
            if np.sum(abs(self.Datas[1][maxArg-1:ruptureArg]) < 0.1*max(abs(self.Datas[1]))):
                ruptureArg = abs(abs(self.Datas[1][maxArg-1:ruptureArg]) - 0.1*max(abs(self.Datas[1]))).argmin() + maxArg-1 -1
            self.Datas[0] = self.Datas[0][:ruptureArg]
            self.Datas[1] = self.Datas[1][:ruptureArg]

            loadboundinf = abs(abs(np.array(self.Datas[1][:maxArg])) - self.stiffboundssearch[0]/100*max(abs(self.Datas[1][:maxArg]))).argmin()
            loadboundsup = abs(abs(np.array(self.Datas[1][:maxArg])) - self.stiffboundssearch[1]/100*max(abs(self.Datas[1][:maxArg]))).argmin()
            argmax, self.deriv = locMax(self.Datas[0][:maxArg], self.Datas[1][:maxArg])

            if not argmax:
                argmax = maxArg
            self.argmax = argmax

            if self.max:
                argmax = len(self.Datas[0][:maxArg])-1
                self.argmax = len(self.Datas[0][:maxArg])-1

            deriv = locMax(self.Datas[0][loadboundinf:loadboundsup], self.Datas[1][loadboundinf:loadboundsup])[1]

            self.stiffBounds = np.array(stiffness(deriv, self.tol)) + loadboundinf
            self.argb1 = self.stiffBounds[0]
            self.argb2 = self.stiffBounds[1]
            self.stiff, b, self.r, _, _ = linregress(self.Datas[0][self.stiffBounds[0]:self.stiffBounds[1]], self.Datas[1][self.stiffBounds[0]:self.stiffBounds[1]])

            if self.extrapol == 1:
                stiff1, b1, _, _, _ = linregress(self.Datas[0][0:10], self.Datas[1][0:10])
                # self.Datas[0] = np.array([-b1/stiff1] + list(self.Datas[0]))
                # self.Datas[1] = np.array([0] + list(self.Datas[1]))

            plt.cla()
            self.areaLine = self.ax.fill_between(self.Datas[0][:argmax], 0, self.Datas[1][:argmax], alpha=0.5, color='blue')
            self.ax.plot(self.Datas[0], self.Datas[1], '-k')
            self.toolbar.update()
            self.maxLine = self.ax.plot(self.Datas[0][argmax], self.Datas[1][argmax], 'or', picker=10)[0]
            self.stiffLine = self.ax.plot(self.Datas[0][self.stiffBounds], [self.stiff*self.Datas[0][self.stiffBounds[0]]+b, self.stiff*self.Datas[0][self.stiffBounds[1]]+b], '-g')[0]
            self.stiffLine1 = self.ax.plot(self.Datas[0][self.stiffBounds[0]], self.Datas[1][self.stiffBounds[0]], 'og', picker=10)[0]
            self.stiffLine2 = self.ax.plot(self.Datas[0][self.stiffBounds[1]], self.Datas[1][self.stiffBounds[1]], 'og', picker=10)[0]
            self.ax.set_title(self.filename)
            self.ax.set_xlabel('Displacement (%s)'%self.disp_unit.split(' ')[0])
            self.ax.set_ylabel('%s (%s)'%(self.columnName, self.col_unit.split(' ')[0]))
            self.canvas.draw()

            self.area = abs(np.trapz(self.Datas[1][:self.argmax], self.Datas[0][:self.argmax]))
            self.stiff = (self.Datas[1][self.stiffBounds[1]] - self.Datas[1][self.stiffBounds[0]]) / \
                         (self.Datas[0][self.stiffBounds[1]] - self.Datas[0][self.stiffBounds[0]])

            self.eventId = []
            self.eventId.append(self.canvas.mpl_connect('button_press_event', self.on_click))
            self.eventId.append(self.canvas.mpl_connect('button_release_event', self.on_release))
            self.eventId.append(self.canvas.mpl_connect('pick_event', self.on_pick))
            self.eventId.append(self.canvas.mpl_connect('motion_notify_event', self.on_motion))

            self.stiffnessval_lab.config(text='%.3f %s/%s'%(self.stiff, self.col_unit, self.disp_unit) +' , r2 = '+'%.5f' %self.r**2)
            self.Fmaxval_lab.config(text='%.3f %s'%(max(abs(self.Datas[1])), self.col_unit))
            self.selectedFmax = max(abs(self.Datas[1]))
            self.selectedFmax_lab.config(text='%.3f %s'%(self.selectedFmax, self.col_unit))
            self.selectedDisp_at_Fmax = self.Datas[0][np.argmax(abs(self.Datas[1]))]
            self.selectedDisp_at_Fmax_lab.config(text='%.3f %s'%(self.selectedDisp_at_Fmax, self.disp_unit))
            self.energyval_lab.config(text='%.4f %s.%s'%(self.area, self.col_unit, self.disp_unit))

        except Exception as e:
            self.loadError = True
            print(e)
            exc_type, exc_obj, exc_tb = sys.exc_info()
            fname = os.path.split(exc_tb.tb_frame.f_code.co_filename)[1]
            print(exc_type, fname, exc_tb.tb_lineno)
            print("Load file ERROR %s" % filepath)

        return

    def update(self):
        """ Update infos and canvas """
        self.areaLine.remove()
        self.areaLine = self.ax.fill_between(self.Datas[0][:self.argmax], 0, self.Datas[1][:self.argmax], alpha=0.5, color='blue')

        self.stiff, b, self.r, _, _ = linregress(self.Datas[0][min(self.argb1, self.argb2):max(self.argb1, self.argb2)], self.Datas[1][min(self.argb1, self.argb2):max(self.argb1, self.argb2)])
        self.stiffLine.set_data(self.Datas[0][[self.argb1, self.argb2]], [self.stiff*self.Datas[0][self.argb1]+b, self.stiff*self.Datas[0][self.argb2]+b])

        self.canvas.draw()

        self.stiffnessval_lab.config(text='%.3f %s/%s'%(self.stiff, self.col_unit, self.disp_unit) +' , r2 = '+'%.5f' %self.r**2)
        self.area = abs(np.trapz(self.Datas[1][:self.argmax], self.Datas[0][:self.argmax]))
        self.selectedFmax = abs(self.Datas[1][self.argmax-1])
        self.selectedFmax_lab.config(text='%.3f %s'%(self.selectedFmax, self.col_unit))
        self.selectedDisp_at_Fmax = abs(self.Datas[0][self.argmax-1])
        self.selectedDisp_at_Fmax_lab.config(text='%.3f %s'%(self.selectedDisp_at_Fmax, self.disp_unit))
        self.energyval_lab.config(text='%.4f %s.%s'%(self.area, self.col_unit, self.disp_unit))
        return

    def stiffChange(self):
        """ Update stiffness and canvas """
        self.stiffBounds = stiffness(self.deriv, float(self.linTol_entry.get()))
        self.stiff, b, self.r, _, _ = linregress(self.Datas[0][self.stiffBounds[0]:self.stiffBounds[1]], self.Datas[1][self.stiffBounds[0]:self.stiffBounds[1]])
        self.stiffLine.set_data(self.Datas[0][self.stiffBounds], [self.stiff*self.Datas[0][self.stiffBounds[0]]+b, self.stiff*self.Datas[0][self.stiffBounds[1]]+b])
        self.stiffLine1.set_data(self.Datas[0][self.stiffBounds[0]], self.Datas[1][self.stiffBounds[0]])
        self.stiffLine2.set_data(self.Datas[0][self.stiffBounds[1]], self.Datas[1][self.stiffBounds[1]])
        self.stiffnessval_lab.config(text='%.3f' %self.stiff+' , r2 = '+'%.5f' %self.r**2)
        self.canvas.draw()

    def on_motion(self, event):
        """
        Mouse motion callback. Drags the selected point and change the
        associated boundary value
        """
        # If mouse pressed and a point is created and no tools from
        # toolbar are selected
        if self.click is True and self.ax.get_navigate_mode() is None:
            # Aspect ratio for better mouse selection
            aspect_ratio = float(self.winfo_height()) / \
                self.winfo_width()

            # Mouse click coordinates
            x = event.xdata
            y = event.ydata
            if not x or not y:
                return

            # Nearest point on curve
            xid = find_nearest([self.Datas[0], self.Datas[1]],
                               [x, y], aspect_ratio)

            if self.thisline == self.maxLine:
                self.argmax = xid+1

            if self.thisline == self.stiffLine1:
                self.argb1 = xid

            if self.thisline == self.stiffLine2:
                self.argb2 = xid

            self.thisline.set_data([self.Datas[0][xid]], [self.Datas[1][xid]])

            # Updates the canvas
            self.update()

    def on_click(self, event):
        self.click = True

    def on_release(self, event):
        """ Mouse button released callback. Deactivates the pressed state """
        self.click = False

    def on_pick(self, event):
        """ Curve selecting callback """
        if self.ax.get_navigate_mode() is None:  # If no tool from toolbar is selected

            self.thisline = event.artist
            # Restores other lines thickness to default
            [plt.setp(line, 'mew', 0.5) for line in self.ax.lines]
            # Increase thickness of selected line
            plt.setp(self.thisline, 'mew', 2.0)

            # Update canavas
            self.canvas.draw()

    def massAnalysis(self, *args, **kwargs):
        filenames = kwargs.get('filenames', None)
        if filenames is None:
            filenames = askopenfilenames(title='Open .xlsx or .csv file from Bose measures',
                                         filetypes=[('Excel file, csv file', '*.xlsx *.csv')])
        path = os.path.split(filenames[0])[0]
        with Workbook(os.path.join(path, 'results.xlsx')) as fileSave:
            worksheet1 = fileSave.add_worksheet()
            worksheet2 = fileSave.add_worksheet()

            worksheet1.write_row(0, 0, ('Filename','Stiffness','r2 (stiffness)','Breakage energy','Selected Fmax','Disp at selected Fmax', 'Real Fmax'))
            for i, filename in enumerate(filenames):
                self.load(filename=filename)
                if self.loadError:
                    worksheet1.write_row(i + 1, 0, (
                    self.filename, "ERROR", "ERROR", "ERROR", "ERROR", "ERROR"))
                    continue
                worksheet1.write_row(i+1, 0, (self.filename,self.stiff,self.r**2,self.area, abs(self.selectedFmax), abs(self.selectedDisp_at_Fmax), max(abs(self.Datas[1]))))
                worksheet2.write(0, i*3, self.filename)
                worksheet2.write(1, i*3, 'Displacement')
                worksheet2.write(1, i*3+1, 'Load')
                worksheet2.write_column(2, i*3, self.Datas[0])
                worksheet2.write_column(2, i*3+1, self.Datas[1])
                plt.savefig(os.path.join(path, 'results_')+self.filename+'.png', dpi=300)

        return

    def controlledMassAnalyzis(self, *args, **kwargs):
        self.filenames = askopenfilenames(title='Open .xlsx or .csv file from Bose measures',
                                          filetypes=[('Excel file, csv file', '*.xlsx *.csv')])
        self.controlled_mass_analysis_check = True
        self.current_file_index = 0
        self.load(filename=self.filenames[self.current_file_index])

    def save(self):
        # with open(os.path.join(self.path, 'results_')+self.filename+'.csv', 'w+') as f:
        #     f.write('Stiffness,'+str(self.stiff)+'\n')
        #     f.write('r2 (stiffness),'+str(self.r**2)+'\n')
        #     f.write('Breakage energy,'+str(self.area)+'\n')
        #     f.write('Fmax,'+str(max(abs(self.Datas[1]))))
        # plt.savefig(os.path.join(self.path, 'results_')+self.filename+'.png', dpi=300)
        with Workbook(os.path.join(self.path, 'results_'+self.filename+'.xlsx')) as fileSave:
            worksheet1 = fileSave.add_worksheet()
            worksheet2 = fileSave.add_worksheet()

            worksheet1.write_row(0, 0, ('Filename','Stiffness','r2 (stiffness)','Breakage energy','Selected Fmax','Disp at selected Fmax', 'Real Fmax'))
            i = 0
            worksheet1.write_row(i+1, 0, (self.filename,self.stiff,self.r**2, self.area, abs(self.selectedFmax), abs(self.selectedDisp_at_Fmax), max(abs(self.Datas[1]))))
            worksheet2.write(0, i*3, self.filename)
            worksheet2.write(1, i*3, 'Displacement')
            worksheet2.write(1, i*3+1, 'Load')
            worksheet2.write_column(2, i*3, self.Datas[0])
            worksheet2.write_column(2, i*3+1, self.Datas[1])
            plt.savefig(os.path.join(self.path, 'results_'+self.filename+'.png'), dpi=300)

        print('%s RESULTS SAVED.'%self.filename)

        if self.controlled_mass_analysis_check:
            if self.current_file_index == 0:
                path = os.path.split(self.filenames[0])[0]
                self.fileSave = Workbook(os.path.join(path, 'results.xlsx'))
                self.worksheet1 = self.fileSave.add_worksheet()
                self.worksheet2 = self.fileSave.add_worksheet()

                self.worksheet1.write_row(0, 0, ('Filename', 'Stiffness', 'r2 (stiffness)', 'Breakage energy', 'Selected Fmax','Disp at selected Fmax', 'Real Fmax'))

            self.worksheet1.write_row(self.current_file_index + 1, 0,
                                 (self.filename, self.stiff, self.r ** 2, self.area, abs(self.selectedFmax), abs(self.selectedDisp_at_Fmax), max(abs(self.Datas[1]))))
            self.worksheet2.write(0, self.current_file_index * 3, self.filename)
            self.worksheet2.write(1, self.current_file_index * 3, 'Displacement')
            self.worksheet2.write(1, self.current_file_index * 3 + 1, 'Load')
            self.worksheet2.write_column(2, self.current_file_index * 3, self.Datas[0])
            self.worksheet2.write_column(2, self.current_file_index * 3 + 1, self.Datas[1])

            self.current_file_index += 1
            if self.current_file_index >= len(self.filenames):
                self.controlled_mass_analysis_check = False
                self.filenames = []
                self.worksheet1 = []
                self.worksheet2 = []
                self.fileSave.close()
                self.fileSave = []
                self.current_file_index = 0
            else:
                self.load(filename=self.filenames[self.current_file_index])

    def _quit(self, *args, **kwargs):
        """ Properly quit the program """
        self.quit()
        self.destroy()



program = MecaAnalysis()
tk.mainloop()
