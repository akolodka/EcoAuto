VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ListEventHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder "eco_transferToWord.Core.Dialogs.Main Transfer.Event Handlers.Controls Handlers"
'@PredeclaredId
Option Explicit
Implements IDialogControlEventHandler

Private WithEvents ListBox As MSForms.ListBox
Attribute ListBox.VB_VarHelpID = -1

Private Type TListEventHandler
    Action As ITransferDialogAction
End Type

Private this As TListEventHandler

Public Function Create( _
       ByVal Action As ITransferDialogAction, _
       ByVal Control As MSForms.Control) As IDialogControlEventHandler
    
    If TypeOf Control Is MSForms.ListBox = False Then _
       Exit Function
    
    With New ListEventHandler
        
        Set .List = Control
        Set .Action = Action
        
        Set Create = .Self
    
    End With

End Function

Public Property Get List() As MSForms.ListBox
    Set List = ListBox
End Property
Public Property Set List(ByVal RHS As MSForms.ListBox)
    Set ListBox = RHS
End Property

Public Property Get Action() As ITransferDialogAction
    Set Action = this.Action
End Property
Public Property Set Action(ByVal RHS As ITransferDialogAction)
    Set this.Action = RHS
End Property

Public Property Get Self() As IDialogControlEventHandler
    Set Self = Me
End Property

'---------------------------------------------
Private Sub ListBox_KeyDown( _
        ByVal KeyCode As MSForms.ReturnInteger, _
        ByVal Shift As Integer)
    
    OnKeyDowned KeyCode

End Sub

Private Sub OnKeyDowned( _
        ByVal KeyCode As MSForms.ReturnInteger)
        
    If KeyCode = vbKeyEscape Then _
       KeyCode = vbEmpty: Action.HideDialog
        
    If KeyCode = vbKeyReturn Then _
       KeyCode = vbEmpty: Action.InitiateTransfer
        
End Sub


