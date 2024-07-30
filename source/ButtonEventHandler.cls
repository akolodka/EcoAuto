VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ButtonEventHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder "eco_transferToWord.Core.Dialogs.Main Transfer.Event Handlers.Controls Handlers"
'@PredeclaredId
Option Explicit
Implements IDialogControlEventHandler

Private WithEvents CommandButton As MSForms.CommandButton
Attribute CommandButton.VB_VarHelpID = -1

Private Type TButtonEventHandler
    Action As ITransferDialogAction
End Type

Private this As TButtonEventHandler

Public Function Create( _
       ByVal Action As ITransferDialogAction, _
       ByVal Control As MSForms.Control) As IDialogControlEventHandler
    
    If TypeOf Control Is MSForms.CommandButton = False Then _
       Exit Function
    
    With New ButtonEventHandler
        
        Set .Button = Control
        Set .Action = Action
        
        Set Create = .Self
    
    End With

End Function

Public Property Get Button() As MSForms.CommandButton
    Set Button = CommandButton
End Property
Public Property Set Button(ByVal RHS As MSForms.CommandButton)
    Set CommandButton = RHS
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
Private Sub CommandButton_KeyDown( _
        ByVal KeyCode As MSForms.ReturnInteger, _
        ByVal Shift As Integer)
    
    OnKeyDowned KeyCode

End Sub

Private Sub OnKeyDowned( _
        ByVal KeyCode As MSForms.ReturnInteger)
        
    If KeyCode = vbKeyEscape Then
       
       KeyCode = vbEmpty
       Action.HideDialog
       
    End If
        
    If KeyCode = vbKeyReturn Then
       
       KeyCode = vbEmpty
       Action.InitiateTransfer
    
    End If
        
End Sub
