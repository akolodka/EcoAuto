VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IEditableExcelWorksheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("eco_transferToWord.Core.Excel.Worksheet")
'@Interface
Option Explicit

Public Sub CloneRowDown(ByVal RowIndex As Long)
End Sub

Public Sub DeleteRow(ByVal RowIndex As Long)
End Sub

Public Property Get Cells(ByVal Row As Long, ByVal Column As Long) As IExcelRange
End Property

Public Property Get Name() As String
End Property

Public Property Let Name(ByVal RHS As String)
End Property
