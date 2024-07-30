VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ComparisonChartsheetHandlerFact"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder "eco_transferToWord.Core.Comparison Results Decoration.Comparison Block Decoration.Chart Sheet Handler"
'@PredeclaredId
Option Explicit
Implements IComparisonChartsheetHandlerFac

Private Type TComparisonChartsheetHandlerFactory
    Chartbook As IExcelWorkbook
End Type

Private this As TComparisonChartsheetHandlerFactory

Public Function Create( _
       ByVal Chartbook As IExcelWorkbook) As IComparisonChartsheetHandlerFac
       
    With New ComparisonChartsheetHandlerFact
        
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
Public Property Get Self() As IComparisonChartsheetHandlerFac
    Set Self = Me
End Property
'---------------------------------------------
Private Function IComparisonChartsheetHandlerFac_Create( _
        ByVal Model As List_IOutputResultRowModel_) As IComparisonChartsheetHandler

    With New ComparisonChartsheetHandler
    
        Set .Chartbook = Chartbook
        Set .Model = Model
    
        Set IComparisonChartsheetHandlerFac_Create = .Self
    
    End With
    
End Function
