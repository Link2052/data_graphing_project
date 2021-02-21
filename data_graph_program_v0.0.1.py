from tkinter import filedialog as F
from tkinter import ttk
from tkinter import *
import time

import os

try:
	from matplotlib import pyplot as graph
except:
	os.system('python3 pip -m install matplotlib')

try:
	import numpy as ny
except:
	os.system('python3 pip -m install numpy')

#from openpyxl import load_workbook
#from openpyxl.utils import get_column_letter

#import numpy
###	a = numpy.asarray([ [1,2,3], [4,5,6], [7,8,9] ])
###	ny.savetxt("foo.csv", a, delimiter=",")

class App(Frame):
	
	def __init__(self, master):
	
		Frame.__init__(self, master, width = 1000, height = 1000)
		self.parent = master
		self.grid()
		self.parent.title("Excel Scrapper")

		self.GUI()
		
	def GUI(self):
	
		bl0 = ttk.Label(self)
		bl0.grid(row = 0, column = 0)
		
		bl1 = ttk.Label(self)
		bl1.grid(row = 4, column = 8)
		
		l0 = ttk.Label(self, text = 'Column Select:')
		l0.grid(row = 2, column = 1)
		
		self.e0 = ttk.Entry(self, width = 80)
		self.e0.grid(row = 1, column = 2, columnspan = 5, sticky = EW)
		
		self.e1 = ttk.Entry(self, width = 80)
		self.e1.grid(row = 2, column = 2, columnspan = 5, sticky = EW)
		
		self.b0 = ttk.Button(self, text = 'File:', command = self.filedial)
		self.b0.grid(row = 1, column = 1, sticky = EW)
		
		self.b1 = ttk.Button(self, text = 'Enter', command = self.process)
		self.b1.grid(row = 3, column = 1, columnspan = 6, sticky = EW)
		
	def filedial(self):
	
		self.filepath = F.askopenfilename(initialdir = 'desktop')
		self.e0.delete(0, 'end')
		self.e0.insert(0, self.filepath)
	
	def process(self):
		
		try:		
			if self.filename:
				pass
		except:
			self.filename = self.e0.get()
		
		self.filesavedirc = os.path.dirname(self.filename)
		
		self.filesavedirc = self.filesavedirc+'/Graph Repository'
			
		try:
			os.makedirs(self.filesavedirc)
		except FileExistsError:
			pass
		
		#wb = openpyxl.load_workbook('origin.xlsx')
		#first_sheet = wb.get_sheet_names()[0]
		#worksheet = wb.get_sheet_by_name(first_sheet)
		
		'''workbook = load_workbook(self.filename)
		sheet = workbook.active
				
		cols = [int(i) for i in self.e1.get().split(',')]
		
		data = [ny.array([item.value for item in sheet[get_column_letter(i)]]) for i in cols]
		
		workbook.close()'''
		__key_words__ = ['time', 'TIME', 'Time']
		
		file = open(self.filename, 'r')
		raw_data = ny.array([line.split(',') for line in file])
		file.close()
		index = ny.where(ny.core.defchararray.find(A, ('TIME', 'time')==0)
		ind = [[], []]
		[[[ind[0].append(item0) for item0 in line[0]], [ind[1].append(item1) for item1 in line[1]]] for line in [ny.where(ny.core.defchararray.find(B, str)==0) for str in __key_words__]]
		__srow__ = ind[0][0]
		__cols__ = [int(i) for i in self.e1.get().split(',')]
		
		labels = [raw_data[__srow__][i] for i in __cols__]
		__cols__.insert(0, ind[1][0])
		proc_data = ny.array([[float(item) for item in line] for line in raw_data[__srow__+1:, __cols__]])
		proc_data = proc_data.T
		
		'''>>> graph.figure(0)
		<Figure size 640x480 with 0 Axes>
		>>> graph.plot(x, y, 'r')
		[<matplotlib.lines.Line2D object at 0x0F76DD48>]
		>>> graph.figrue(1)
		Traceback (most recent call last):
		  File "<stdin>", line 1, in <module>
		AttributeError: module 'matplotlib.pyplot' has no attribute 'figrue'
		>>> graph.figure(1)
		<Figure size 640x480 with 1 Axes>
		>>> graph.plot(x, z)
		[<matplotlib.lines.Line2D object at 0x0F76DE98>]
		>>> graph.figure(0)
		<Figure size 640x480 with 1 Axes>
		>>> graph.plot(x, rr)
		[<matplotlib.lines.Line2D object at 0x0F77E058>]
		>>> graph.show()'''
		self.fig_group = []
		for i in range(len(__cols__)-1):
			self.fig_group.append(graph.figure(i))
			current_figure = self.fig_group[-1].add_subplot(111)
			current_figure.plot(proc_data[0], proc_data[i])
			graph.grid()
			
			current_figure.set_title("Thermocouple data "+labels[i])
			current_figure.set_xlabel("Time [min]")
			current_figure.set_ylabel("Temp. ["+u'\N{DEGREE SIGN}'+"C]")
			
			time_str = time.strftime(%Y_%m_%d-%H_%M_%S)
			
			self.fig_group[-1].savefig(self.filesavedirc+'/'+labels[i]+'_'+time_str+'.png', format = 'png')
		
		
		
app = App(Tk())
app.mainloop()
