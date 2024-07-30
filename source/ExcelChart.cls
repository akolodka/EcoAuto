VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ExcelChart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder "eco_transferToWord.Core.Excel.Chart"
'@PredeclaredId
Option Explicit
Implements IExcelChart

Private Type TChartObject
    Chart As Excel.ChartObject
End Type

Private this As TChartObject

Public Function Create(ByVal Chart As Excel.ChartObject)
    
    With New ExcelChart
    
        Set .Chart = Chart
        Set Create = .Self
        
    End With
    
End Function

Public Property Get Chart() As Excel.ChartObject
    Set Chart = this.Chart
End Property
Public Property Set Chart(ByVal RHS As Excel.ChartObject)
    Set this.Chart = RHS
End Property

Public Property Get Self() As IExcelChart
    Set Self = Me
End Property

Private Sub CopyAsPicture()
    Chart.CopyPicture
End Sub

Private Sub ApplyScale( _
        ByVal MinValue As Double, _
        ByVal MaxValue As Double)
        
    Chart.Activate
    
    ActiveChart.Axes(xlValue).MinimumScale = MinValue
    ActiveChart.Axes(xlValue).MaximumScale = MaxValue

End Sub

'---------------------------------------------
Private Sub IExcelChart_CopyAsPicture()
    CopyAsPicture
End Sub

Private Sub IExcelChart_ApplyScale(ByVal MinValue As Double, ByVal MaxValue As Double)
    ApplyScale MinValue, MaxValue
End Sub
