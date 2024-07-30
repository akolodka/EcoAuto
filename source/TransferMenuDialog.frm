VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} TransferMenuDialog 
   ClientHeight    =   5865
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5055
   OleObjectBlob   =   "TransferMenuDialog.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "TransferMenuDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder "eco_transferToWord.Core.Dialogs.Main Transfer.Dialog"
Option Explicit

Implements IDialogView

Public Event BoxContentChanged(ByVal Text As String, ByVal Label As MSForms.Label)
Public Event ParticipantsFilterApplied(ByVal Text As String)
Public Event CancelRequested()
Public Event BoxExited()

Public Event TransferInitiated()
Public Event HelpRequested()

Private Type TTransferMenuDialog
    isCancelled As Boolean
End Type

Private this As TTransferMenuDialog


Private Sub UserForm_Activate()
    Me.TourBox.SetFocus
End Sub

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

Private Sub TourBox_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

    Const ASCII_LEFT As Long = 48
    Const ASCII_RIGHT As Long = 57
    
    If KeyAscii < ASCII_LEFT Or KeyAscii > ASCII_RIGHT Then _
        KeyAscii = vbEmpty

End Sub
Private Sub TourBox_Change()

    RaiseEvent BoxContentChanged(TourBox.Text, TourNumberHintLabel)
    RaiseEvent ParticipantsFilterApplied(TourBox.Text)
    
End Sub

Private Sub TourBox_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    RaiseEvent BoxExited
End Sub

'---------------------------------------------
Private Sub FactoryNumberBox_Change()
    RaiseEvent BoxContentChanged(FactoryNumberBox.Text, FactoryNumberHintLabel)
End Sub
Private Sub FactoryNumberBox_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    RaiseEvent BoxExited
End Sub

'---------------------------------------------
Private Sub RespondentBox_Change()
    RaiseEvent BoxContentChanged(RespondentBox.Text, RespondentHintLabel)
End Sub
Private Sub RespondentBox_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    RaiseEvent BoxExited
End Sub

'---------------------------------------------
Private Sub HelpButton_Click()
    RaiseEvent HelpRequested
End Sub

Private Sub TransferButton_MouseUp( _
        ByVal Button As Integer, _
        ByVal Shift As Integer, _
        ByVal X As Single, _
        ByVal Y As Single)
    
    RaiseEvent TransferInitiated
    
End Sub

Private Sub CancelButton_Click()
    RaiseEvent CancelRequested
End Sub

'---------------------------------------------
Private Function IDialogView_Show() As Boolean

    Me.Show vbModal
    IDialogView_Show = Not this.isCancelled

End Function

Private Sub IDialogView_Hide()
    OnCancel
End Sub


