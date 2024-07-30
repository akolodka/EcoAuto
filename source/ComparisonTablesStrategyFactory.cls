VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ComparisonTablesStrategyFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "eco_transferToWord.Core.Comparison Results Decoration.Comparison Block Decoration.Strategies.Tables"
Option Explicit

Implements IComparisonTemplateFileStrategy
Implements IComparisonProgressNotifyStrate
Implements IComparisonBlockStrategyFactory

Const ProgressFooterText As String = "Таблицы сличений"
Const FileMask As String = "comparisonTables"
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
    
    With New ComparisonTablesStrategy
        
        Set .TableHandler = DecorationTableHandler.Create(Document)
        Set IComparisonBlockStrategyFactory_Create = .Self
        
    End With
    
End Function
