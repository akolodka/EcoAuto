VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} TransferHelpDialog 
   ClientHeight    =   6705
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5250
   OleObjectBlob   =   "TransferHelpDialog.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "TransferHelpDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder "eco_transferToWord.Core.Dialogs.Help.Dialog"
Option Explicit
Implements IDialogView

Public Event CloseButtonPressed()

Private Type TTransferHelpDialog
    isCancelled As Boolean
End Type

Private this As TTransferHelpDialog

Private Sub UserForm_QueryClose( _
        ByRef Cancel As Integer, _
        ByRef CloseMode As Integer)
    
    If CloseMode = VbQueryClose.vbFormControlMenu Then
    
        Cancel = True
        OnCancel
        
    End If
    
End Sub

Private Sub OnCancel()
    
    this.isCancelled = True
    Me.Hide
        
End Sub

'---------------------------------------------
Private Sub CloseButton_Click()
    RaiseEvent CloseButtonPressed
End Sub

Private Sub CloseButton_KeyDown( _
        ByVal KeyCode As MSForms.ReturnInteger, _
        ByVal Shift As Integer)
    
    If KeyCode = vbKeyEscape Then _
       RaiseEvent CloseButtonPressed
    
End Sub

'---------------------------------------------
Private Function IDialogView_Show() As Boolean
    
    Me.Show vbModal
    IDialogView_Show = Not this.isCancelled

End Function

Private Sub IDialogView_Hide()
    OnCancel
End Sub


