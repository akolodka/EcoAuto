Attribute VB_Name = "Program"
'@Folder("eco_transferToWord")
Option Explicit

Public Sub Main()

    Dim Initial As IInitializationService
    Set Initial = InitializationService.Create()
    
    Dim Validator As IReferenceValuesValidator
    Set Validator = ReferenceValuesValidator.Create(Initial)
    
    If (Validator.IsReferenceDataUnique = False) Then
        
        Validator.SuggestCorrection
        Exit Sub
        
    End If
    
    Dim Dialog As ITransferDialogService
    Set Dialog = TransferMenuPresenter.Create(Initial)
    
    Dialog.Show
    
End Sub

Public Sub Help()
    
    Dim Dialog As ITransferDialogService
    Set Dialog = TransferHelpPresenter.Create()
    
    Dialog.Show
    
End Sub
