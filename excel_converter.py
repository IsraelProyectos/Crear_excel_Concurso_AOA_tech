#!/usr/bin/env python
#coding=utf-8

import wx
# from openpyxl import *
# from openpyxl.styles import Color, PatternFill, Font, Border, colors, borders, Side
import readAndSaveExcel

class MyFrame(wx.Frame):
            def __init__(self, parent, title):
                  wx.Frame.__init__(self, parent, title=title, size=(400,200))

                  self.pathFile = ''
                  self.txtRuta = wx.TextCtrl(self, pos=(10,50), size=(250,20), style=wx.TE_READONLY)
                  self.buttonFind = wx.Button(self, label="Buscar...", pos=(270,50), size=(100,20))
                  self.buttonFind.Bind(wx.EVT_BUTTON, self.openFile)
                  self.buttonExecute = wx.Button(self, label="Convertir", pos=(270,100), size=(100,20))
                  
                  self.buttonExecute.Disable()
                  self.labelEstadoOperacion= wx.StaticText(self, pos=(10,130), size=(360,20), style=wx.TE_READONLY)
                  # self.labelEstadoOperacion.SetBackgroundColour( wx.Colour( 255, 255, 255))
                  self.labelEstadoOperacion.SetForegroundColour(wx.Colour(255, 0, 0))
                  self.buttonExecute.Bind(wx.EVT_BUTTON, self.createExcel)
                  

                  self.columna_excel = [ ]
                  self.todas_columnas = [ ]
                  self.registro_excel_final = [ ]
                  self.registros_excel_final =[ ]
                  self.fields = ['EMAIL', 'CODIGO_1', 'CODIGO_2', 'CODIGO_3', 'CODIGO_4', 'CODIGO_5', 'NOMBRE_CODIGO_1',
                  						  'NOMBRE_CODIGO_2', 'NOMBRE_CODIGO_3', 'NOMBRE_CODIGO_4', 'NOMBRE_CODIGO_5', 	
										  'OBJ_AO_1', 'OBJ_AO_2', 'OBJ_AO_3', 'OBJ_AO_4', 'OBJ_AO_5', 
										  'OBJ_AOA_1', 'OBJ_AOA_2',	'OBJ_AOA_3', 'OBJ_AOA_4', 'OBJ_AOA_5',	
										  'EMAIL_CC', 'EMAIL_REMITENTE', 'EMAIL_CONTACTO', 'NOMBRE']
                  self.email = ""
                  self.i = 0
                  self.z = 2
                  self.columnaDelExcel = 1
                  self.Centre(True)
                  self.SetBackgroundColour(wx.Colour( 252, 255, 228))
                  self.Show(True)

            def createExcel(self, e):
            	self.rase = readAndSaveExcel.ReadAndSaveExcell(self.pathFile)
            	result = self.rase.readExcel(self)
            	mensajeOk = self.rase.colorMensaje
            	if mensajeOk:
            		self.labelEstadoOperacion.SetForegroundColour(wx.Colour( 0, 138, 0))
            		self.labelEstadoOperacion.SetLabel(result)
            	else:
            		self.labelEstadoOperacion.SetLabel(result)



            
            def openFile(self, e):
				try:
					self.labelEstadoOperacion.SetLabel("")

	            	#Creando la ventana para escoger el archivo. Solo para archivos con extension .xlsx(Archivo excel 2010)
					with wx.FileDialog(self, "Abrir archivo .xlsx", wildcard="XLSX files (*.xlsx)|*.xlsx",
						style=wx.FD_OPEN | wx.FD_FILE_MUST_EXIST) as fileDialog:

						#Cerrar ventana de diálogo al dar a cancelar
						if fileDialog.ShowModal() == wx.ID_CANCEL:
							return

				        #Guardando el path del archivo en variable
				        self.pathFile = fileDialog.GetPath()

				        #Asignando ese path al textBox
				        try:
				            self.txtRuta.SetValue(self.pathFile)
				            self.buttonExecute.Enable()
				        except IOError:
							wx.LogError("Cannot open file '%s'." % newfile)
					
				except KeyError:
					print(err)

app = wx.App(False)
frame = MyFrame(None, 'Creación Fichero Excel')
app.MainLoop()
