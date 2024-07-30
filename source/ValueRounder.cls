VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ValueRounder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder "eco_transferToWord.Core.Initialization.Comparison Mapping.Block Data Mapper.Value Round"
'@PredeclaredId
Option Explicit
Implements IValueRounder

Private Type TValueRounder
    
    Value As String
    Uncertainty As String
    
End Type

Private this As TValueRounder

Public Function Create( _
       ByVal AbsoluteValue As String, _
       Optional ByVal AbsoluteUncertainty As String) As IValueRounder

    With New ValueRounder
        
        .Value = AbsoluteValue
        .Uncertainty = AbsoluteUncertainty
        
        Set Create = .Self
        
    End With

End Function

Public Property Get Value() As String
    Value = this.Value
End Property
Public Property Let Value(ByVal RHS As String)
    this.Value = RHS
End Property

Public Property Get Uncertainty() As String
    Uncertainty = this.Uncertainty
End Property
Public Property Let Uncertainty(ByVal RHS As String)
    this.Uncertainty = RHS
End Property

Public Property Get Self() As IValueRounder
    Set Self = Me
End Property

Public Function GetValue() As String
    
    GetValue = Value
    
    If (IsValueRoundable = False) Then _
        Exit Function
    
    Dim MultiplyCoefficient As Long
    MultiplyCoefficient = GetMultiplyCoefficient(Uncertainty)
    
    Dim mValue As Long
    mValue = CDbl(Value) * MultiplyCoefficient
    
    Dim mUncertainty As Long
    mUncertainty = CDbl(Uncertainty) * MultiplyCoefficient
    
    Dim FirstValueNumber As Long
    FirstValueNumber = Left(mUncertainty, 1)
    
    Dim RoundedValue As Long
    
    If FirstValueNumber < 4 Then
        RoundedValue = RoundValue(mValue, -Len(CStr(mUncertainty)) + 2)
    Else
        RoundedValue = RoundValue(mValue, -Len(CStr(mUncertainty)) + 1)
    End If
    
    Dim Result As Double
    Result = RoundedValue / MultiplyCoefficient
    
    Dim ResultFormat As String
    ResultFormat = GetUncertaintyFormat(GetUncertainty())
    
    GetValue = Format(Result, ResultFormat)
    
End Function
    Private Function GetUncertaintyFormat( _
            ByVal Uncertainty As String) As String
                    
        If (IsNumeric(Uncertainty) = False) Then _
           Exit Function
                    
        Const ReplaceSymbol As String = "123456789"
            
        Dim WorkUncertainty As String
        WorkUncertainty = Uncertainty
            
        Dim Predicate As Boolean
        Predicate = (Len(Uncertainty) = 1) _
                    And (Uncertainty >= 1) _
                    And (Uncertainty < 4)
            
        If (Predicate = True) Then _
           WorkUncertainty = Uncertainty & ",0"
                    
        Dim Result As String
                    
        Dim i As Long
        For i = 1 To Len(WorkUncertainty)
                        
            Dim Sym As String
            Sym = Mid(WorkUncertainty, i, 1)
                                               
            Dim ReplacedSym As String
            ReplacedSym = IIf(InStr(ReplaceSymbol, Sym) > vbEmpty, _
                              vbEmpty, Sym)
                
            Result = Result & ReplacedSym
                            
        Next
                    
        Dim FormattedResult As String
        FormattedResult = Replace(Result, ",", ".")
                    
        GetUncertaintyFormat = FormattedResult
                    
    End Function
    Private Function RoundValue( _
            ByVal Value As Long, _
            ByVal DigitNumber As Long) As Long
        
        Dim Result As Long
        Result = Application.WorksheetFunction.Round(Value, DigitNumber)
        
        RoundValue = Result
        
    End Function
    Private Function IsValueRoundable() As Boolean
        
        Dim Predicate As Boolean
        Predicate = IsNumeric(Value) And IsNumeric(Uncertainty)
        
        If (Predicate = False) Then _
            Exit Function
            
        Predicate = (CDbl(Value) > vbEmpty) And (CDbl(Uncertainty) > vbEmpty)
        
        IsValueRoundable = Predicate
        
    End Function
    Private Function GetMultiplyCoefficient( _
            ByVal Value As Double) As Long
        
        Dim Coefficient As Long
        Coefficient = 1
        
        Dim MultipliedValue As Long
        MultipliedValue = Value * Coefficient
        
        Do
            If (MultipliedValue >= 10) Then _
                Exit Do
            
            Coefficient = Coefficient * 10
            MultipliedValue = Value * Coefficient
            
        Loop
        
        GetMultiplyCoefficient = Coefficient
        
    End Function
Public Function GetUncertainty() As String
    
    GetUncertainty = Uncertainty
    
    If (IsUncertaintyRoundable = False) Then _
        Exit Function
    
    Dim MultiplyCoefficient As Long
    MultiplyCoefficient = GetMultiplyCoefficient(Uncertainty)
    
    Dim mUncertainty As Long
    mUncertainty = CDbl(Uncertainty) * MultiplyCoefficient
    
    Dim FirstValueNumber As Long
    FirstValueNumber = Left(mUncertainty, 1)
    
    Dim RoundedUncertainty As Long
    
    If FirstValueNumber < 4 Then
        RoundedUncertainty = RoundValue(mUncertainty, -Len(CStr(mUncertainty)) + 2)
    Else
        RoundedUncertainty = RoundValue(mUncertainty, -Len(CStr(mUncertainty)) + 1)
    End If
    
    Dim Result As Double
    Result = RoundedUncertainty / MultiplyCoefficient
    
    Dim ResultFormat As String
    ResultFormat = GetUncertaintyFormat(Result)
    
    GetUncertainty = Format(Result, ResultFormat)

End Function
    Private Function IsUncertaintyRoundable() As Boolean
    
        Dim Predicate As Boolean
        Predicate = IsNumeric(Uncertainty)
        
        If (Predicate = False) Then _
            Exit Function
            
        Predicate = (CDbl(Uncertainty) > vbEmpty)
       
        IsUncertaintyRoundable = Predicate
        
    End Function
Public Function GetRoundedToHundredths() As String
    
    GetRoundedToHundredths = Value
    
    If IsRoundableToHundredths = False _
        Then Exit Function
    
    Dim Result As Double
    Result = Application.WorksheetFunction.Round(Value, 2)
    
    GetRoundedToHundredths = Result
    
End Function
    Private Function IsRoundableToHundredths() As Boolean
        
        Dim Predicate As Boolean
        Predicate = IsNumeric(Value)
               
        IsRoundableToHundredths = Predicate
        
    End Function
'---------------------------------------------
Private Function IValueRounder_GetValue() As String
    IValueRounder_GetValue = GetValue
End Function
Private Function IValueRounder_GetUncertainty() As String
    IValueRounder_GetUncertainty = GetUncertainty
End Function
Private Function IValueRounder_GetRoundedToHundredths() As String
    IValueRounder_GetRoundedToHundredths = GetRoundedToHundredths
End Function


