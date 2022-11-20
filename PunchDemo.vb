Option Strict Off
Option Explicit On
Imports System.Text

<System.Runtime.InteropServices.ProgId("Class1_NET.Class1")> Public Class LiaisonExcel
	Public Declare Function GetPrivateProfileString Lib "kernel32" (ByVal lpAppName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As StringBuilder, ByVal nSize As Integer, ByVal lpFileName As String) As Integer
	Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As StringBuilder, ByVal lpString As StringBuilder, ByVal lpFileName As String) As Long

	Public xlApp As Microsoft.Office.Interop.Excel.Application
	Public xlBook As Microsoft.Office.Interop.Excel.Workbook
	Public xlRang As Microsoft.Office.Interop.Excel.Range
	Public xlSheet As Microsoft.Office.Interop.Excel.Worksheet
	Public Col As Short
	Public Fichier As String
	Public ini As String
	Public Function Init() As String
		Fichier = $"{Environment.CurrentDirectory()}\Punch_Demo.xlsx"
		ini = $"{Environment.CurrentDirectory()}\Punch_Demo.ini"
		Return Fichier
	End Function
	Public Function CreerEntetes() As Boolean
		xlApp = New Microsoft.Office.Interop.Excel.Application
		With xlApp
			xlBook = .Workbooks.Open(Filename:=Fichier, ReadOnly:=False, Editable:=True)
			xlSheet = xlBook.Worksheets(1)
		End With
		With xlSheet
			Const V As String = "Utilisateur"
			.Name = V
			.Cells._Default(1, 1) = "Type"
			.Cells._Default(1, 2) = "Capacitée"
			.Cells._Default(1, 3) = "Date"
			.Cells._Default(1, 4) = "Heure"
		End With
		xlApp.Visible = True
		xlApp.SaveWorkspace(xlSheet.Name)
		xlRang = Nothing
		xlSheet = Nothing
		xlBook = Nothing
		xlApp = Nothing
		Return True
	End Function
	Public Function EcrireCellule(ByRef Fichier As String, ByRef Nom As String, ByRef Ligne As Short, ByRef Colonne As Short, ByRef Texte As String)
		xlApp = New Microsoft.Office.Interop.Excel.Application
		With xlApp
			xlBook = .Workbooks.Open(Filename:=Fichier, ReadOnly:=False, Editable:=True)
			xlSheet = xlBook.Worksheets(1)
		End With
		With xlSheet
			.Name = Nom
			.Cells._Default(Ligne, Colonne) = Texte
		End With
		xlApp.Visible = True
		xlApp.SaveWorkspace(xlSheet.Name)
		xlRang = Nothing
		xlSheet = Nothing
		xlBook = Nothing
		xlApp = Nothing
	End Function
	Public Function LireCellule(ByRef Fichier As String, ByRef Nom As String, ByRef Ligne As Short, ByRef Colonne As Short, ByRef Texte As String) As String
		Dim tmp As String
		xlApp = New Microsoft.Office.Interop.Excel.Application
		With xlApp
			xlBook = .Workbooks.Open(Filename:=Fichier, ReadOnly:=False, Editable:=True)
			xlSheet = xlBook.Worksheets(1)
			With xlSheet
				.Name = Nom
				tmp = .Cells._Default(Ligne, Colonne)
			End With
			xlApp.SaveWorkspace((xlSheet.Name))
			xlApp.Quit()
			xlRang = Nothing
			xlSheet = Nothing
			xlBook = Nothing
			xlApp = Nothing
		End With
		Return tmp
	End Function
End Class