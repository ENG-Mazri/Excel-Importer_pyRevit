#Mazri's BIM Diary
#eng.mazri@gmail.com
import sys
import os
import clr
import math
import rpw
import Autodesk
from Autodesk.Revit.UI import *
from Autodesk.Revit.DB import *
import Autodesk.Revit.UI.Selection
from Autodesk.Revit.DB import Transaction 
from rpw.ui.forms import (Console,TaskDialog)
clr.AddReference('System.Windows.Forms')
clr.AddReference('System.Drawing')
clr.AddReference('IronPython.Wpf')
#import wpf creator and base window
import wpf
from System import Windows
from System.Windows import MessageBox
from System.Windows.Forms import*
from pyrevit import UI
from pyrevit import script
from rpw.ui.forms import (Console, FlexForm, Label, ComboBox, TextBox, TextBox, CheckBox, Separator, Button, Alert, CommandLink, TaskDialog, select_file)
#find the path of ui xaml
xamlfile = script.get_bundle_file('ui.xaml')

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
t = Transaction(doc, 'Task Dialog') 
#Import Excel stuff
clr.AddReference('Microsoft.Office.Interop.Excel, Version=11.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c')
from Microsoft.Office.Interop import Excel
from System.Runtime.InteropServices import Marshal
form_components = select_file()

# UI form binding
class MyWindow(Windows.Window):
	def __init__(self):
		wpf.LoadComponent(self, xamlfile)
		
	@property
	def tx_shtN(self):
		return self.textbox_shtN.Text
	@property
	def tx_shn1(self):
		return self.textbox_shn1.Text
	@property
	def tx_shn2(self):
		return self.textbox_shn2.Text
	@property
	def tx_shnm1(self):
		return self.textbox_shnm1.Text
	@property
	def tx_shnm2(self):
		return self.textbox_shnm2.Text
		

	def bu1_click(self, sender , args):
		
		sheetNumber = float(self.tx_shtN)
		sheet_name1 = self.tx_shn1
		sheet_name2 = self.tx_shn2
		sheet_num1 = self.tx_shnm1
		sheet_num2 = self.tx_shnm2
		#UI.TaskDialog.Show("Excel Test" , sheet_name2 )

		xl_file = script.get_bundle_file(form_components)
		ex = Excel.ApplicationClass()
		ex.Visible = False
		ex.DisplayAlerts = False
		# Workbook
		workbook = ex.Workbooks.Open(xl_file)
		# WorkSheet
		ws = workbook.Worksheets(sheetNumber)
		# Cell range
		x1range = ws.Range[sheet_num1 , sheet_num2]
		r1 = x1range.Value2
		x2range = ws.Range[sheet_name1 , sheet_name2]
		r2 = x2range.Value2

		lst1 = []
		for i in r1:
			lst1.append(i)

		lst2 = []
		for j in r2:
			lst2.append(j)

     			
		t.Start()
	
		sht_names = lst2
		sht_numbers = lst1
		sht_list = []

		for num in range(len(sht_numbers)):
			sht = ViewSheet.Create(doc , ElementId(4428))
			sht.Name = sht_names[num]
			sht.SheetNumber = sht_numbers[num]
			sht_list.append(sht)

		#ex.ActiveWorkbook.Close(False)
		#Marshal.ReleaseComObject(ws)
		#Marshal.ReleaseComObject(workbook)
		#Marshal.ReleaseComObject(ex)	
		t.Commit()
		UI.TaskDialog.Show("Excel Test" , "Sheets were created successfully" )


MyWindow().ShowDialog()

'''		if self.check_box1 == True:
			#self.textbox1.Text = "i'm indeed a great textbox"	
			MessageBox.Show("ListBox doing well")		
		else:
			#self.textbox1.Text = "i'm NOT a great textbox"
			MessageBox.Show("ListBox doing well too")'''