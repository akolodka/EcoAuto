VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ComparisonChartsHandlerFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder "eco_transferToWord.Core.Comparison Results Decoration.Comparison Block Decoration.Chart Handler"
'@PredeclaredId
Option Explicit
Implements IComparisonChartsHandlerFactory

Private Type TComparisonChartsHandlerFactory
    Chartbook As IExcelWorkbook
End Type

Private this As TComparisonChartsHandlerFactory

Public Function Create( _
       ByVal Chartbook As IExcelWorkbook) As IComparisonChartsHandlerFactory
       
    With New ComparisonChartsHandlerFactory
        
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

Public Property Get Self() As IComparisonChartsHandlerFactory
    Set Self = Me
End Property
'---------------------------------------------
Private Function IComparisonChartsHandlerFactory_Create( _
        ByVal BlockIndex As Long, _
        ByVal ParticipantName As String) As IComparisonChartsHandler
    
    With New ComparisonChartsHandler
        
        Set .ChartSheet = Chartbook.Worksheets(BlockIndex)
        .ParticipantName = ParticipantName
        
        Set IComparisonChartsHandlerFactory_Create = .Self
    
    End With

End Function

