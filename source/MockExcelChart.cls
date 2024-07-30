VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MockExcelChart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@PredeclaredId
'@Folder("Tests.Mocks.Excel.Chart")
Option Explicit
Implements IExcelChart

Public Function Create() As IExcelChart
    
    With New MockExcelChart
        Set Create = .Self
    End With
    
End Function

Public Property Get Self() As IExcelChart
    Set Self = Me
End Property

Private Sub CopyAsPicture()

Debug.Print
Debug.Print "Mock ChartObject.CopyAsPicture"

End Sub

Private Sub ApplyScale( _
        ByVal MinValue As Double, _
        ByVal MaxValue As Double)


Debug.Print
Debug.Print "Mock ChartObject.ApplyScale " & _
            "   MinValue: "; MinValue; _
            "   MaxValue: "; MaxValue

End Sub
'---------------------------------------------
Private Sub IExcelChart_ApplyScale(ByVal MinValue As Double, ByVal MaxValue As Double)
    ApplyScale MinValue, MaxValue
End Sub

Private Sub IExcelChart_CopyAsPicture()
    CopyAsPicture
End Sub
