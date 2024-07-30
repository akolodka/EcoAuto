VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MockWordTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder("Tests.Mocks.Word.Table")
'@PredeclaredId
Option Explicit

Implements IWordTable
Implements IFormattableWordTable
Implements IAppendableWordTable

Public Function Create() As IFormattableWordTable
    
    With New MockWordTable
        Set Create = .Self
    End With
    
End Function

Public Property Get Self() As IFormattableWordTable
    Set Self = Me
End Property
'---------------------------------------------
Private Property Get IFormattableWordTable_RowsCount() As Long
    IFormattableWordTable_RowsCount = 2
End Property

Private Sub IFormattableWordTable_ApplyFont( _
        ByVal FontName As String)
    
Debug.Print
Debug.Print "Mock WordTable.ApplyFont"
    
End Sub

'---------------------------------------------
Private Sub IAppendableWordTable_AddRow()
    
Debug.Print
Debug.Print "Mock WordTable.AddRow"
    
End Sub

Private Sub IAppendableWordTable_AppendValue( _
        ByVal Column As Long, _
        ByVal Value As String, _
        Optional ByVal isBold As Boolean)
    
Debug.Print
Debug.Print "Mock WordTable.AddRows" & vbNewLine & _
            "   Column: "; Column; " " & vbNewLine & _
            "   Value: "; Value; " " & vbNewLine & _
            "   isBold: "; isBold
    
End Sub

Private Property Get IAppendableWordTable_IsLastRowFilled() As Boolean
    
    Static Result As Boolean
    Result = Not Result
    
Debug.Print
Debug.Print "Mock WordTable.IsLastRowFilled" & vbNewLine & _
            "   Result: "; Result

    IAppendableWordTable_IsLastRowFilled = Result
    
End Property

Private Sub IAppendableWordTable_RemoveDublicateRows()
    
Debug.Print
Debug.Print "Mock WordTable.RemoveDublicateRows"

End Sub

Private Sub IFormattableWordTable_UnderlineHeader()

Debug.Print
Debug.Print "Mock WordTable.UnderlineHeader"

End Sub

Private Sub IFormattableWordTable_MergeEqualCellsByRows(ByVal TargetColumn As Long)

Debug.Print
Debug.Print "Mock WordTable.MergeEqualCellsByRows" & vbNewLine & _
            "   Column: "; TargetColumn
End Sub

Private Sub IFormattableWordTable_SetRowsHeight(ByVal HeightInCentimeters As Double)

Debug.Print
Debug.Print "Mock WordTable.SetRowsHeight" & vbNewLine & _
            "   Height: "; HeightInCentimeters
End Sub

Private Property Get IWordTable_IsFilled() As Boolean

    Static Result As Boolean
    Result = Not Result

Debug.Print
Debug.Print "Mock WordTable.IsFilled" & vbNewLine & _
            "   Result: "; Result

    IWordTable_IsFilled = Result

End Property
