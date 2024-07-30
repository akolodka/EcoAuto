VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ControlsEventHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder "eco_transferToWord.Core.Dialogs.Main Transfer.Event Handlers"
'@PredeclaredId
Option Explicit
Implements ITransferMenuEventHandler

Private Type TControlsEventHandler

    View As TransferMenuDialog
    Action As ITransferDialogAction
    
    KeyDownControls As List_IDialogControlEventHandle_
    
End Type

Private this As TControlsEventHandler

Public Function Create( _
       ByVal View As TransferMenuDialog, _
       ByVal Action As ITransferDialogAction) As ITransferMenuEventHandler
    
    With New ControlsEventHandler
        
        Set .View = View
        Set .Action = Action
        
        Set Create = .Self
    
    End With
    
End Function

Public Property Get View() As TransferMenuDialog
    Set View = this.View
End Property
Public Property Set View(ByVal RHS As TransferMenuDialog)
    Set this.View = RHS
End Property

Public Property Get Action() As ITransferDialogAction
    Set Action = this.Action
End Property
Public Property Set Action(ByVal RHS As ITransferDialogAction)
    Set this.Action = RHS
End Property

Public Property Get KeyDownControls() As List_IDialogControlEventHandle_
    Set KeyDownControls = this.KeyDownControls
End Property
Public Property Set KeyDownControls(ByVal RHS As List_IDialogControlEventHandle_)
    Set this.KeyDownControls = RHS
End Property

Public Property Get Self() As ITransferMenuEventHandler

    GetBindedControls
    
    Set Self = Me
    
End Property

Private Sub GetBindedControls()
        
    Dim uncheckedHandlers As List_IDialogControlEventHandle_
    Set uncheckedHandlers = GetUncheckedHandlers()

    Set KeyDownControls = GetCheckedHandlers(uncheckedHandlers)
    
End Sub

Private Function GetUncheckedHandlers() As List_IDialogControlEventHandle_
            
    Dim Result As List_IDialogControlEventHandle_
    Set Result = New List_IDialogControlEventHandle_
            
    Dim Control As MSForms.Control
    For Each Control In View.Controls
                                
        Result.Add TextBoxEventHandler.Create(Action, Control, View)
        Result.Add ButtonEventHandler.Create(Action, Control)
        Result.Add ListEventHandler.Create(Action, Control)
        Result.Add ComboxAsListEventHandler.Create(Action, Control)
                    
    Next
            
    Set GetUncheckedHandlers = Result

End Function

Private Function GetCheckedHandlers( _
        ByVal unchecked As List_IDialogControlEventHandle_) As List_IDialogControlEventHandle_
            
    Dim Result As List_IDialogControlEventHandle_
    Set Result = New List_IDialogControlEventHandle_
            
    Dim i As Long
    For i = 1 To unchecked.Count
                
        If (IsCorrect(unchecked(i)) = True) Then _
           Result.Add unchecked(i)
                
    Next
            
    Set GetCheckedHandlers = Result
            
End Function

Private Function IsCorrect( _
        ByVal Handler As IDialogControlEventHandler) As Boolean
                
    IsCorrect = (Handler Is Nothing = False)
                
End Function


