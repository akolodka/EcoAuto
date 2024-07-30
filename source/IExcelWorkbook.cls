VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IExcelWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("eco_transferToWord.Core.Excel.Workbook")
'@Interface
Option Explicit

Public Property Get Worksheets(ByVal Name As String) As IExcelWorksheet
End Property

Public Property Get LastSheet() As IExcelWorksheet
End Property

Public Property Get FirstSheet() As IExcelWorksheet
End Property

Public Function IsSheetExist(ByVal Name As String) As Boolean
End Function
