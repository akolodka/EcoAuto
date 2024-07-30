VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "KeyValuePair"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder "eco_transferToWord.Core.Document IO.KeyValue Pair"
'@PredeclaredId
Option Explicit

Implements IKeyValuePair

Private Type TKeyValuePair
    
    Key As ITransferKey
    Value As ITransferValue
    
End Type

Private this As TKeyValuePair

Public Function Create( _
       ByVal unsafeKey As String, _
       ByVal unsafeValue As Variant) As IKeyValuePair
    
    With New KeyValuePair
    
        Set .Key = TransferKey.Create(unsafeKey)
        Set .Value = TransferValue.Create(unsafeValue)
        
        Set Create = .Self
    
    End With
    
End Function

Public Property Get Key() As ITransferKey
    Set Key = this.Key
End Property
Public Property Set Key(ByVal RHS As ITransferKey)
    Set this.Key = RHS
End Property


Public Property Get Value() As ITransferValue
Attribute Value.VB_UserMemId = 0
    Set Value = this.Value
End Property
Public Property Set Value(ByVal RHS As ITransferValue)
    Set this.Value = RHS
End Property

Public Property Get Self() As IKeyValuePair
    Set Self = Me
End Property

'---------------------------------------------
Private Property Get IKeyValuePair_Key() As ITransferKey
    Set IKeyValuePair_Key = Key
End Property
'@DefaultMember
Private Property Get IKeyValuePair_Value() As ITransferValue
    Set IKeyValuePair_Value = Value
End Property
