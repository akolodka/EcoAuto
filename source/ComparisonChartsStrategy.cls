VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ComparisonChartsStrategy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "eco_transferToWord.Core.Comparison Results Decoration.Comparison Block Decoration.Strategies.Charts"
Option Explicit
Implements IComparisonBlockDecorationStrat

Private Type TComparisonChartsStrategy

    Writer As IDocumentWriter
    
    ChartHandlerFactory As IComparisonChartsHandlerFactory
    SheetHandlerFactory As IComparisonChartsheetHandlerFac
    
End Type

Private this As TComparisonChartsStrategy

Public Property Get Self() As IComparisonBlockDecorationStrat
    Set Self = Me
End Property

Public Property Get Writer() As IDocumentWriter
    Set Writer = this.Writer
End Property
Public Property Set Writer(ByVal RHS As IDocumentWriter)
    Set this.Writer = RHS
End Property

Public Property Get ChartHandlerFactory() As IComparisonChartsHandlerFactory
    Set ChartHandlerFactory = this.ChartHandlerFactory
End Property
Public Property Set ChartHandlerFactory(ByVal RHS As IComparisonChartsHandlerFactory)
    Set this.ChartHandlerFactory = RHS
End Property

Public Property Get SheetHandlerFactory() As IComparisonChartsheetHandlerFac
    Set SheetHandlerFactory = this.SheetHandlerFactory
End Property
Public Property Set SheetHandlerFactory(ByVal RHS As IComparisonChartsheetHandlerFac)
    Set this.SheetHandlerFactory = RHS
End Property

Public Sub FillBlock( _
       ByVal Model As List_IOutputResultRowModel_, _
       ByVal ParticipantNumber As String)
    
    Dim Handler As IComparisonChartsheetHandler
    Set Handler = SheetHandlerFactory.Create(Model)
    
    If (Handler.IsSheetFilled = False) Then _
        Handler.FillSheet
        
    Dim i As Long
    For i = 1 To Model.Count
        
        If (Model(i).ParticipantNumber = ParticipantNumber) Then _
            WriteChart Model(i)

    Next

End Sub
    Private Sub WriteChart(ByVal Model As IComparisonChartIdentifyModel)
        
        Dim Handler As IComparisonChartsHandler
        Set Handler = ChartHandlerFactory.Create(Model.BlockIndex, Model.ParticipantName)
        
        Dim Chart As IKeyValuePair
        Set Chart = Handler.GetChart()
        
        Writer.InsertIntoDocument Chart, ReplaceFirstEntry
        
    End Sub
'---------------------------------------------
Private Sub IComparisonBlockDecorationStrat_FillBlock(ByVal Model As List_IOutputResultRowModel_, ByVal ParticipantNumber As String)
    FillBlock Model, ParticipantNumber
End Sub


