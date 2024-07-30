VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TransferKey"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder "eco_transferToWord.Core.Document IO.Transfer KeyValue.Key"
'@PredeclaredId
Option Explicit
Implements ITransferKey

Private Type TTransferKey
    Value As String
End Type

Private this As TTransferKey

Public Function Create( _
       ByVal unsafeKey As String) As ITransferKey
    
    With New TransferKey
        
        .Value = GetBracedValue(unsafeKey)
        Set Create = .Self
        
    End With
    
End Function

Private Function GetBracedValue( _
        ByVal unsafeKey As String) As String
        
    Const OpenSymbol As String = "{"
    Const CloseSymbol As String = "}"
        
    If unsafeKey = vbNullString Then _
       Exit Function
        
    Dim predicateOne As Boolean
    predicateOne = (Left(unsafeKey, 1) = OpenSymbol)
        
    Dim predicateTwo As Boolean
    predicateTwo = (Right(unsafeKey, 1) = CloseSymbol)
        
    If predicateOne And predicateTwo Then _
       Exit Function
        
    GetBracedValue = OpenSymbol & unsafeKey & CloseSymbol
    
End Function

Public Property Get Self() As ITransferKey
    Set Self = Me
End Property
Public Property Get Value() As String
Attribute Value.VB_UserMemId = 0
    Value = this.Value
End Property
Public Property Let Value(ByVal RHS As String)
    this.Value = RHS
End Property
'---------------------------------------------
'@DefaultMember
Private Property Get ITransferKey_Value() As String
    ITransferKey_Value = Value
End Property
