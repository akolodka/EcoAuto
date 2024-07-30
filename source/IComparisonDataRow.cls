VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IComparisonDataRow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("eco_transferToWord.Core.Excel.Range")
'@Interface
Option Explicit

Public Function GetValue(ByVal Column As Long) As String
End Function

Public Property Get ColumnsCount() As Long
End Property
