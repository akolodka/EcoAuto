VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ProgressDialog 
   Caption         =   "Выполнение..."
   ClientHeight    =   825
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   3975
   OleObjectBlob   =   "ProgressDialog.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "ProgressDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder "eco_transferToWord.Core.Progress Bar.Dialog"
Option Explicit
Implements IDialogView

Private Type TView
    isCancelled As Boolean
End Type

Private this As TView

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
Private Function IDialogView_Show() As Boolean
    
    Me.Show vbModeless
    IDialogView_Show = Not this.isCancelled
    
End Function

Private Sub IDialogView_Hide()
    OnCancel
End Sub


