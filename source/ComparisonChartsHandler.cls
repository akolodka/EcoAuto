VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ComparisonChartsHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "eco_transferToWord.Core.Comparison Results Decoration.Comparison Block Decoration.Chart Handler"
Option Explicit
Implements IComparisonChartsHandler

Private Type TComparisonChartsHandler

    ChartSheet As IChartableExcelWorksheet
    ParticipantName As String

End Type

Private this As TComparisonChartsHandler

Public Property Get Self() As IComparisonChartsHandler
    Set Self = Me
End Property

Public Property Get ChartSheet() As IChartableExcelWorksheet
    Set ChartSheet = this.ChartSheet
End Property
Public Property Set ChartSheet(ByVal RHS As IChartableExcelWorksheet)
    Set this.ChartSheet = RHS
End Property

Public Property Get ParticipantName() As String
    ParticipantName = this.ParticipantName
End Property
Public Property Let ParticipantName(ByVal RHS As String)
    this.ParticipantName = RHS
End Property

Public Function GetChart() As IKeyValuePair
    
    If (ChartSheet Is Nothing = True) Then _
        Exit Function
    
    WriteParticipantName ChartSheet
    Set GetChart = GetComparisonChart
    
End Function
Private Sub WriteParticipantName( _
        ByVal Sheet As IEditableExcelWorksheet)
        
    Sheet.Cells(4, 13).Value = ParticipantName
    
End Sub
Private Function GetComparisonChart() As IKeyValuePair
    
        Const KeyName As String = "picture"
        
        Dim ChartObject As IExcelChart
        Set ChartObject = ChartSheet.GetChart()
        
        Dim Result As IKeyValuePair
        Set Result = KeyValuePair.Create(KeyName, ChartObject)
        
        Set GetComparisonChart = Result
    
    End Function
'---------------------------------------------
Private Function IComparisonChartsHandler_GetChart() As IKeyValuePair
    Set IComparisonChartsHandler_GetChart = GetChart
End Function


