VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TransferValue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder "eco_transferToWord.Core.Document IO.Transfer KeyValue.Value"
'@PredeclaredId
Option Explicit
Implements ITransferValue

Private Type TTransferValue
    Value As Variant
End Type

Private this As TTransferValue

Public Function Create( _
       ByVal unsafeValue As Variant) As ITransferValue
    
    If (IsTransferValue(unsafeValue) = True) Then
    
        Set Create = unsafeValue
        Exit Function
    
    End If
    
    With New TransferValue
    
        If (IsObject(unsafeValue) = True) Then
            Set .Value = unsafeValue
        Else
            .Value = GetSafeTextValue(unsafeValue)
        End If
        
        Set Create = .Self
    
    End With

End Function

Private Function IsTransferValue( _
        ByVal Value As Variant) As Boolean
        
    If (IsObject(Value) = False) Then _
       Exit Function
            
    If (Value Is Nothing = True) Then _
       Exit Function
        
    If TypeOf Value Is ITransferValue Then _
       IsTransferValue = True
        
End Function

Private Function GetSafeTextValue( _
        ByVal Value As String) As String
        
    If Value = vbNullString Then _
       Exit Function
        
    Dim Result As String
    Result = Trim$(Value)
        
    Do While (Predicate(Result) = True)
        Result = Left$(Result, Len(Result) - 1)
    Loop
        
    GetSafeTextValue = Result
        
End Function

Private Function Predicate( _
        ByVal Value As String) As Boolean
            
    Dim rightChar As String
    rightChar = Right$(Value, 1)
            
    'Пустые символы имеют значения ASCII от 0 до 32
    Const BlankSymAscii As Long = 32
            
    Predicate = (Asc(rightChar) <= BlankSymAscii)
        
End Function

Public Property Get Self() As ITransferValue
    Set Self = Me
End Property

'@DefaultMember
Public Property Get Value() As Variant
Attribute Value.VB_UserMemId = 0

    If IsObject(this.Value) Then
        Set Value = this.Value
    Else
        Value = this.Value
    End If
    
End Property

Public Property Let Value(ByVal RHS As Variant)
    this.Value = RHS
End Property
Public Property Set Value(ByVal RHS As Variant)
    Set this.Value = RHS
End Property

'---------------------------------------------
Private Property Get ITransferValue_Value() As Variant
    
    If IsObject(this.Value) Then
        Set ITransferValue_Value = this.Value
    Else
        ITransferValue_Value = this.Value
    End If
    
End Property
