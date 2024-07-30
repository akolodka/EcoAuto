VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WordTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder("eco_transferToWord.Core.Word.Table")
'@PredeclaredId
Option Explicit

Implements IWordTable
Implements IFormattableWordTable
Implements IAppendableWordTable

Private Type TWordTable
    Table As Word.Table
End Type

Private this As TWordTable

Public Function Create( _
       ByVal Table As Word.Table) As IWordTable
    
    With New WordTable
    
        Set .Table = Table
        Set Create = .Self
        
    End With
    
End Function

Public Property Get Table() As Word.Table
    Set Table = this.Table
End Property
Public Property Set Table(ByVal RHS As Word.Table)
    Set this.Table = RHS
End Property

Public Property Get Self() As IWordTable
    Set Self = Me
End Property

Private Sub AddRow()
    Table.Rows.Add
End Sub

Private Sub AppendValue( _
        ByVal Column As Long, _
        ByVal Value As String, _
        Optional ByVal isBold As Boolean)
    
    If (Column > Table.Columns.Count) Then _
       Exit Sub
    
    With Table.Cell(Table.Rows.Count, Column).Range
            
        .Text = Value
        .Bold = isBold
        
    End With
    
End Sub

Private Function IsLastRowFilled() As Boolean
    
    Dim isCellFree As Boolean
    
    Dim i As Long
    For i = 1 To Table.Columns.Count
        
        Dim cellText As String
        cellText = Table.Cell(Table.Rows.Count, i).Range.Text
        
        'Символ абзаца в таблице занимает 2 единицы длины
        Const FreeCellLengh As Long = 2
        
        If Len(cellText) = FreeCellLengh Then
        
            isCellFree = True
            Exit For
            
        End If
        
    Next
    
    IsLastRowFilled = (isCellFree = False)
    
End Function

Private Sub UnderlineHeader()
    Table.Rows(1).Borders(wdBorderBottom).LineStyle = wdLineStyleDouble
End Sub

Private Sub MergeEqualCellsByRows( _
        ByVal Column As Long)
    
    Dim mergedCells As List_IMergedCell_
    Set mergedCells = GetMergedCells(Column)
    
    If mergedCells.Count = vbEmpty Then _
       Exit Sub
    
    Dim i As Long
    For i = 1 To mergedCells.Count
    
        Table.Cell(mergedCells(i).StartRow, Column).Merge _
        Table.Cell(mergedCells(i).EndRow, Column)
                    
        Dim cleanedValue As String
        cleanedValue = GetCleanedCellValue(mergedCells(i).Value)
                    
        Table.Cell(mergedCells(i).StartRow, Column).Range.Text = cleanedValue
        
    Next
    
    Table.Rows.DistributeHeight

End Sub
    Private Function GetMergedCells( _
            ByVal Column As Long) As List_IMergedCell_
            
        Const StartRow As Long = 2
            
        Dim Result As List_IMergedCell_
        Set Result = New List_IMergedCell_
            
        Dim cellText As String
        cellText = Table.Cell(StartRow, Column).Range.Text
            
        Dim merged As IMergedCell
        Set merged = MergedCell.Create( _
                     StartRow, cellText)
        
        Dim i As Long
        For i = StartRow To Table.Rows.Count
                
            cellText = Table.Cell(i, Column).Range.Text
                
            If merged.Value = cellText Then
                merged.EndRow = i
                
            Else
                Result.Add merged
                    
                Set merged = MergedCell.Create( _
                             i, cellText)
                
            End If
        
        Next
            
        Result.Add merged
            
        Set GetMergedCells = Result
                
    End Function
Private Function GetCleanedCellValue( _
        ByVal uncleanedCellValue) As String
        
    GetCleanedCellValue = Left(uncleanedCellValue, Len(uncleanedCellValue) - 2)
        
End Function

Private Sub RemoveDublicateRows()
    
    Dim StartRow As Long
    StartRow = Table.Rows.Count
    
    Const EndRow As Long = 2
    
    Dim Reference As List_String_
    Set Reference = GetTableRowData(StartRow)
    
    Dim Compared As List_String_
    
    Dim i As Long
    For i = StartRow - 1 To EndRow Step -1
            
        Set Compared = GetTableRowData(i)
        
        If (Compared.EqualsTo(Reference) = True) Then
            Table.Rows(i).Delete
        Else
            Set Reference = Compared
        End If
        
    Next
    
End Sub

Private Function GetTableRowData( _
        ByVal RowIndex As Long) As List_String_
        
    If RowIndex > Table.Rows.Count Then _
       Exit Function
        
    Dim Result As List_String_
    Set Result = New List_String_
        
    Dim i As Long
    For i = 1 To Table.Columns.Count
        
        Dim Value As String
        Value = Table.Cell(RowIndex, i).Range.Text
            
        Result.Add Value
        
    Next
        
    Set GetTableRowData = Result

End Function
Private Sub SetRowsHeight( _
        ByVal HeightInCentimeters As Double)
    
    'Application.CentimetersToPoints требует WordApp
    Const PointsInCentimeter As Double = 28.35
    
    Table.Rows.SetHeight HeightInCentimeters * PointsInCentimeter, wdRowHeightAtLeast

End Sub
Private Sub ApplyFont( _
        ByVal FontName As String)
    
    Table.Range.Font.Name = FontName

End Sub
'---------------------------------------------
Private Property Get IWordTable_IsFilled() As Boolean
    IWordTable_IsFilled = Table.Rows.Count > 1
End Property
'---------------------------------------------
Private Sub IFormattableWordTable_ApplyFont(ByVal FontName As String)
    ApplyFont FontName
End Sub
Private Sub IFormattableWordTable_UnderlineHeader()
    UnderlineHeader
End Sub

Private Sub IFormattableWordTable_MergeEqualCellsByRows(ByVal TargetColumn As Long)
    MergeEqualCellsByRows TargetColumn
End Sub

Private Sub IFormattableWordTable_SetRowsHeight(ByVal HeightInCentimeters As Double)
    SetRowsHeight HeightInCentimeters
End Sub
'---------------------------------------------
Private Sub IAppendableWordTable_AddRow()
    AddRow
End Sub
Private Sub IAppendableWordTable_AppendValue(ByVal Column As Long, ByVal Value As String, Optional ByVal isBold As Boolean)
    AppendValue Column, Value, isBold
End Sub
Private Property Get IAppendableWordTable_IsLastRowFilled() As Boolean
    IAppendableWordTable_IsLastRowFilled = IsLastRowFilled
End Property

