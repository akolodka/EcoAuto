VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IExcelWorkbooksService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("eco_transferToWord.Core.Excel")
'@Interface
Option Explicit

Public Function GetFrom(ByVal WorkbookPath As String) As IExcelWorkbook
End Function

Public Function OpenFrom(ByVal WorkbookPath As String) As IExcelWorkbook
End Function

Public Property Get ActiveBook() As IExcelWorkbook
End Property




