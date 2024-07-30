VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AverageCalculator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "eco_transferToWord.Core.Reference Values Validator.Average Calculator"
Option Explicit
Implements IAverageCalculator

Private Type TAverageCalculator

    ValueSum As Double
    Count As Long
    
End Type

Private this As TAverageCalculator

Public Property Get ValueSum() As Double
    ValueSum = this.ValueSum
End Property
Public Property Let ValueSum(ByVal RHS As Double)
    this.ValueSum = RHS
End Property

Public Property Get Count() As Long
    Count = this.Count
End Property
Public Property Let Count(ByVal RHS As Long)
    this.Count = RHS
End Property

Public Sub Append(ByVal Value As Double)

    ValueSum = ValueSum + Value
    Count = Count + 1
    
End Sub
Public Function GetAverage() As Double
    
    If (Count = vbEmpty) Then _
        Exit Function
        
    GetAverage = ValueSum / Count
End Function
'---------------------------------------------
Private Sub IAverageCalculator_Append(ByVal Value As Double)
    Append Value
End Sub
Private Function IAverageCalculator_GetAverage() As Double
    IAverageCalculator_GetAverage = GetAverage
End Function

