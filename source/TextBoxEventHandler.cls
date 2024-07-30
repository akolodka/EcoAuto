VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TextBoxEventHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder "eco_transferToWord.Core.Dialogs.Main Transfer.Event Handlers.Controls Handlers"
'@PredeclaredId
Option Explicit
Implements IDialogControlEventHandler

Private WithEvents TextBox As MSForms.TextBox
Attribute TextBox.VB_VarHelpID = -1
Private WithEvents DialogView As TransferMenuDialog
Attribute DialogView.VB_VarHelpID = -1

Private Type TTextBoxEventHandler

    Action As ITransferDialogAction
    Text As String
    
End Type

Private this As TTextBoxEventHandler

Public Function Create( _
       ByVal Action As ITransferDialogAction, _
       ByVal Control As MSForms.Control, _
       ByVal View As TransferMenuDialog) As IDialogControlEventHandler
    
    If TypeOf Control Is MSForms.TextBox = False Then _
       Exit Function
    
    With New TextBoxEventHandler
        
        Set .Action = Action
        Set .Box = Control
        Set .View = View

        Set Create = .Self
    
    End With

End Function

Public Property Get View() As TransferMenuDialog
    Set View = DialogView
End Property
Public Property Set View(ByVal RHS As TransferMenuDialog)
    Set DialogView = RHS
End Property

Public Property Get Action() As ITransferDialogAction
    Set Action = this.Action
End Property
Public Property Set Action(ByVal RHS As ITransferDialogAction)
    Set this.Action = RHS
End Property

Public Property Get Box() As MSForms.TextBox
    Set Box = TextBox
End Property
Public Property Set Box(ByVal RHS As MSForms.TextBox)
    Set TextBox = RHS
End Property

Public Property Get Text() As String
    Text = this.Text
End Property
Public Property Let Text(ByVal RHS As String)
    this.Text = RHS
End Property

Public Property Get Self() As IDialogControlEventHandler
    Set Self = Me
End Property

'---------------------------------------------
Private Sub TextBox_KeyDown( _
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

'---------------------------------------------
Private Sub TextBox_MouseUp( _
        ByVal Button As Integer, _
        ByVal Shift As Integer, _
        ByVal X As Single, _
        ByVal Y As Single)
    
    If Box.Text = Text Then _
       Exit Sub
    
    Box.SelStart = vbEmpty
    Box.SelLength = Len(Box.Text)
    
    Text = Box.Text
    
End Sub

'---------------------------------------------
Private Sub DialogView_BoxExited()
    Text = vbNullString
End Sub


