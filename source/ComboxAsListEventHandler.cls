VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ComboxAsListEventHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder("eco_transferToWord.Core.Dialogs.Main Transfer.Event Handlers.Controls Handlers")
'@PredeclaredId
Option Explicit
Implements IDialogControlEventHandler

Private WithEvents ComboBox As MSForms.ComboBox
Attribute ComboBox.VB_VarHelpID = -1

Private Type TComBoxEventHandler
    Action As ITransferDialogAction
End Type

Private this As TComBoxEventHandler

Public Function Create( _
       ByVal Action As ITransferDialogAction, _
       ByVal Control As MSForms.Control) As IDialogControlEventHandler
    
    If TypeOf Control Is MSForms.ComboBox = False Then _
       Exit Function
    
    With New ComboxAsListEventHandler
        
        Set .ComBox = Control
        Set .Action = Action
        
        Set Create = .Self
    
    End With

End Function

Public Property Get ComBox() As MSForms.ComboBox
    Set ComBox = ComboBox
End Property
Public Property Set ComBox(ByVal RHS As MSForms.ComboBox)
    Set ComboBox = RHS
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
Private Sub ComboBox_KeyDown( _
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

