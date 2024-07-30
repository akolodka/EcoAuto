VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ComparisonChartsStrategyFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@PredeclaredId
'@Folder "eco_transferToWord.Core.Comparison Results Decoration.Comparison Block Decoration.Strategies.Charts"
Option Explicit

Implements IComparisonTemplateFileStrategy
Implements IComparisonProgressNotifyStrate
Implements IComparisonBlockStrategyFactory

Const ProgressFooterText As String = "Диаграммы сличений"
Const FileMask As String = "comparisonCharts"

Private Type TComparisonChartsStrategyFactory
    Chartbook As IExcelWorkbook
End Type

Private this As TComparisonChartsStrategyFactory
 
Public Function Create( _
       ByVal Chartbook As IExcelWorkbook) As IComparisonBlockStrategyFactory
    
    With New ComparisonChartsStrategyFactory
        
        Set .Chartbook = Chartbook
        Set Create = .Self
        
    End With
    
End Function

Public Property Get Chartbook() As IExcelWorkbook
    Set Chartbook = this.Chartbook
End Property
Public Property Set Chartbook(ByVal RHS As IExcelWorkbook)
    Set this.Chartbook = RHS
End Property

Public Property Get Self() As IComparisonBlockStrategyFactory
    Set Self = Me
End Property
'---------------------------------------------
Private Property Get IComparisonTemplateFileStrategy_TemplateFileMask() As String
    IComparisonTemplateFileStrategy_TemplateFileMask = FileMask
End Property
'---------------------------------------------
Private Property Get IComparisonProgressNotifyStrate_ProgressFooter() As String
    IComparisonProgressNotifyStrate_ProgressFooter = ProgressFooterText
End Property
'---------------------------------------------
Private Function IComparisonBlockStrategyFactory_Create( _
        ByVal Document As IWordDocument, _
        ByVal Writer As IDocumentWriter) As IComparisonBlockDecorationStrat
    
    With New ComparisonChartsStrategy
        
        Set .Writer = Writer
        
        Set .ChartHandlerFactory = ComparisonChartsHandlerFactory.Create(Chartbook)
        Set .SheetHandlerFactory = ComparisonChartsheetHandlerFact.Create(Chartbook)
        
        Set IComparisonBlockStrategyFactory_Create = .Self
        
    End With
    
End Function

