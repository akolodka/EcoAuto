VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MockExcelRange"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder("Tests.Mocks.Excel.Range")
'@PredeclaredId
Option Explicit

Implements IExcelRange
Implements IMarkableExcelRange
Implements IComparisonDataExcelRange
Implements IComparisonDataRow

Private Type TMockExcelRange

    Value As String
    RowIndex As Long
    
End Type

Private this As TMockExcelRange

Public Function Create( _
       Optional ByVal RowIndex As Long = vbEmpty) As IExcelRange
       
    With New MockExcelRange
        
        .Value = "Mocked RangeValue"
        .RowIndex = RowIndex
        
        Set Create = .Self
        
    End With

End Function

Public Property Get RowIndex() As Long
    RowIndex = this.RowIndex
End Property
Public Property Let RowIndex(ByVal RHS As Long)
    this.RowIndex = RHS
End Property

Public Property Get Value() As String
    Value = this.Value
End Property
Public Property Let Value(ByVal RHS As String)
    this.Value = RHS
End Property

Public Property Get Self() As IExcelRange
    Set Self = Me
End Property

Public Sub Calculate()

Debug.Print
Debug.Print "Mock ExcelRange.Calculate"

End Sub

Public Sub Mark()

Debug.Print
Debug.Print "Mock ExcelRange.Mark"

End Sub

Public Function GetRow( _
       ByVal Index As Long) As IExcelRange
    
    Dim Result As IExcelRange
    Set Result = MockExcelRange.Create(Index)
    
    Set GetRow = Result
    
End Function

Public Property Get RowsCount() As Long
    
    Const Result As Long = 7
    
    RowsCount = Result
    
End Property


Private Function GetValue( _
        ByVal Column As Long) As String
    
    Dim Result As String
    
    Select Case Column
        
        Case 2: Result = "Mock ParticipantNumber"
        Case 3: Result = "Mock ParticipantTour"
        Case 5: Result = "Mock ParticipantType"
        Case 6: Result = "Mock ParticipantOrganisation"
        Case Else: Result = "Mock ComparisonBlock Value from Column = " & Column
    
    End Select
        
    GetValue = Result
    
End Function

Private Property Get ColumnsCount() As Long

    Const Result As Long = 10
    
    ColumnsCount = Result

End Property
'---------------------------------------------
Private Property Get IExcelRange_Value() As String
    IExcelRange_Value = Value
End Property
Private Property Let IExcelRange_Value(ByVal RHS As String)
    Value = RHS
End Property
Private Sub IExcelRange_Calculate()
    Calculate
End Sub
'---------------------------------------------
Private Sub IMarkableExcelRange_Mark()
    Mark
End Sub
'---------------------------------------------
Private Function IComparisonDataExcelRange_GetRow(ByVal Index As Long) As IExcelRange
    Set IComparisonDataExcelRange_GetRow = GetRow(Index)
End Function
Private Property Get IComparisonDataExcelRange_RowsCount() As Long
    IComparisonDataExcelRange_RowsCount = RowsCount
End Property
'---------------------------------------------
Private Function IComparisonDataRow_GetValue(ByVal Column As Long) As String
    IComparisonDataRow_GetValue = GetValue(Column)
End Function
Private Property Get IComparisonDataRow_ColumnsCount() As Long
    IComparisonDataRow_ColumnsCount = ColumnsCount
End Property


