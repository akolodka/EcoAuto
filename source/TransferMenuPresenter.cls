VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TransferMenuPresenter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder "eco_transferToWord.Core.Dialogs.Main Transfer"
'@PredeclaredId
Option Explicit

Implements ITransferDialogService
Implements ITransferDialogAction

Private WithEvents DialogView As TransferMenuDialog
Attribute DialogView.VB_VarHelpID = -1

Private Type TTransferMenuPresenter

    Subcontracts As List_IFile_
    ParticipantsFiller As IParticipantListFiller
    
    ValidationStatus As IValidationStatus

    SubcontractHandler As ISubcontractHandler
    Cache As ITransferMenuCacheHandler
    
    HelpDialog As ITransferDialogService
    KeyDowned As ITransferMenuEventHandler
    
    Transfer As ITransferService
    ControlObjectHandler As IControlObjectHandler
    
End Type

Private this As TTransferMenuPresenter

Public Function Create( _
       ByVal Initial As ITransferDialogInitialization) As ITransferDialogService

    With New TransferMenuPresenter
        
        Set .View = New TransferMenuDialog
        Set .Cache = TransferMenuCacheHandler.Create(Initial.FileSystemObject)
        
        Set .Subcontracts = Initial.SubcontractFiles
        
        Set .ParticipantsFiller = ParticipantListFiller.Create( _
                                  .View.ParticipantList, Initial.Participants)

        Set .ValidationStatus = Initial.ValidationStatus
        
        Set .SubcontractHandler = Initial.SubcontractHandler
        Set .ControlObjectHandler = Initial.ControlObjectHandler
        
        Set .KeyDowned = ControlsEventHandler.Create(.View, .Self)
        
        Set .HelpDialog = TransferHelpPresenter.Create(.Self)
        Set .Transfer = Initial.TransferService
        
        Set Create = .Self
        
    End With
    
End Function

Public Property Get Subcontracts() As List_IFile_
    Set Subcontracts = this.Subcontracts
End Property
Public Property Set Subcontracts(ByVal RHS As List_IFile_)
    Set this.Subcontracts = RHS
End Property

Public Property Get ParticipantsFiller() As IParticipantListFiller
    Set ParticipantsFiller = this.ParticipantsFiller
End Property
Public Property Set ParticipantsFiller(ByVal RHS As IParticipantListFiller)
    Set this.ParticipantsFiller = RHS
End Property

Public Property Get ValidationStatus() As IValidationStatus
    Set ValidationStatus = this.ValidationStatus
End Property
Public Property Set ValidationStatus(ByVal RHS As IValidationStatus)
    Set this.ValidationStatus = RHS
End Property

Public Property Get SubcontractHandler() As ISubcontractHandler
    Set SubcontractHandler = this.SubcontractHandler
End Property
Public Property Set SubcontractHandler(ByVal RHS As ISubcontractHandler)
    Set this.SubcontractHandler = RHS
End Property

Public Property Get Cache() As ITransferMenuCacheHandler
    Set Cache = this.Cache
End Property
Public Property Set Cache(ByVal RHS As ITransferMenuCacheHandler)
    Set this.Cache = RHS
End Property

Public Property Get HelpDialog() As ITransferDialogService
    Set HelpDialog = this.HelpDialog
End Property
Public Property Set HelpDialog(ByVal RHS As ITransferDialogService)
    Set this.HelpDialog = RHS
End Property

Public Property Get ControlObjectHandler() As IControlObjectHandler
    Set ControlObjectHandler = this.ControlObjectHandler
End Property
Public Property Set ControlObjectHandler(ByVal RHS As IControlObjectHandler)
    Set this.ControlObjectHandler = RHS
End Property

Public Property Get View() As TransferMenuDialog
    Set View = DialogView
End Property
Public Property Set View(ByVal RHS As TransferMenuDialog)
    Set DialogView = RHS
End Property

Public Property Get KeyDowned() As ITransferMenuEventHandler
    Set KeyDowned = this.KeyDowned
End Property
Public Property Set KeyDowned(ByVal RHS As ITransferMenuEventHandler)
    Set this.KeyDowned = RHS
End Property

Public Property Get Transfer() As ITransferService
    Set Transfer = this.Transfer
End Property
Public Property Set Transfer(ByVal RHS As ITransferService)
    Set this.Transfer = RHS
End Property

Public Property Get Self() As ITransferDialogService
    Set Self = Me
End Property

Private Property Get Dialog() As IDialogView
    Set Dialog = View
End Property

'---------------------------------------------
Private Sub ITransferDialogService_Show()

    View.RespondentBox = Cache.Respondent
    View.FactoryNumberBox = Cache.ControlObjectFactoryNumber

    FillSubcontracts
    SelectSubcontracts
    
    FillConrolObjectsNumbers
    
    FillParticipants
    FillValidationStatus
    
    Dialog.Show
    
End Sub
    Private Sub FillSubcontracts()
            
        Dim Filler As IPresenterDelegate
        Set Filler = SubcontractListFiller.Create( _
                     View.SubcontractList, Subcontracts)
                    
        Filler.Execute
                
    End Sub
    Private Sub SelectSubcontracts()
            
        Dim Selector As IPresenterDelegate
        Set Selector = SubcontractListSelector.Create( _
                       View.SubcontractList, Cache.SelectedSubcontract)
            
        Selector.Execute
    
    End Sub
    Private Sub FillConrolObjectsNumbers()
        
        Dim i As Long
        For i = 1 To ControlObjectHandler.FactoryNumbers.Count
            
            View.FactoryNumberBox.AddItem _
                ControlObjectHandler.FactoryNumbers(i)
            
        Next
        
    End Sub
    Private Sub FillParticipants( _
            Optional ByVal FilterKey As String)
            
        Dim Filler As IPresenterDelegate
        Set Filler = ParticipantsFiller.Create(FilterKey)
                    
        Filler.Execute
            
    End Sub
    Private Sub FillValidationStatus()
                    
        View.ValidationStatusLabel.Caption = ValidationStatus.Message
        View.ValidationStatusLabel.ForeColor = ValidationStatus.Color
              
        If (ValidationStatus.IsTransferAvailable = False) Then _
           View.TransferButton.Caption = View.CancelButton.Caption
            
    End Sub
'---------------------------------------------
Private Sub DialogView_BoxContentChanged( _
        ByVal Text As String, _
        ByVal Label As MSForms.Label)
    
    Label.Visible = (Text = vbNullString)
    FillValidationStatus
    
End Sub

'---------------------------------------------
Private Sub DialogView_HelpRequested()

    OnDialogHide
    HelpDialog.Show

End Sub

Private Sub OnDialogHide()
    
    Dialog.Hide
        
    Cache.Respondent = View.RespondentBox
    Cache.ControlObjectFactoryNumber = View.FactoryNumberBox
        
    Set Cache.SelectedSubcontract = SubcontractListSelector.GetSelectedIndexes(View.SubcontractList)
        
    Cache.Save
        
End Sub

'---------------------------------------------
Private Sub DialogView_ParticipantsFilterApplied( _
        ByVal Text As String)
    
    FillParticipants Text

End Sub
Private Sub DialogView_CancelRequested()
    OnDialogHide
End Sub
Private Sub DialogView_TransferInitiated()
    OnTransferInitiated
End Sub
Private Sub OnTransferInitiated()
    
    If (ValidationStatus.IsTransferAvailable = False) Then
    
        OnDialogHide
        Exit Sub
        
    End If
    
    If View.ParticipantList.ListCount = vbEmpty Then
        
        NotifyIncorrectBoxInput View.TourBox, "Требуется хотя бы один участник."
        Exit Sub
            
    End If
    
    If View.FactoryNumberBox.Text = vbNullString Then
        
        NotifyIncorrectBoxInput View.FactoryNumberBox, "Введите номер объекта для контроля."
        Exit Sub
            
    End If
    
     If View.RespondentBox.Text = vbNullString Then
        
        NotifyIncorrectBoxInput View.RespondentBox, "Укажите ответственное лицо."
        Exit Sub
            
    End If
    
    OnDialogHide
        
    If (ValidationStatus.IsTransferAvailable = False) Then _
       Exit Sub
        
    InitiateTransfer
        
End Sub
    Private Sub NotifyIncorrectBoxInput( _
            ByVal Box As MSForms.Control, _
            ByVal Message As String)
            
        View.ValidationStatusLabel.Caption = Message
        View.ValidationStatusLabel.ForeColor = vbRed
            
        Dim SourceColor As String
        SourceColor = Box.BorderColor
                
        Box.BorderColor = vbRed
        Box.SelStart = vbEmpty
        
        Box.SelLength = Len(Box.Text)
        Box.SetFocus
            
        Dim TimeStamp As Double
        TimeStamp = Timer
        
        Const TimeDelay As Double = 0.5
            
        Do While Timer - TimeStamp < TimeDelay
            DoEvents
        Loop
            
        Box.BorderColor = SourceColor
        View.Repaint

    End Sub
    Private Sub InitiateTransfer()
    
        Dim Model As TransferMenuModelDTO
        Set Model = New TransferMenuModelDTO
         
        Model.Respondent = View.RespondentBox
        Model.Subcontract = GetSubcontractValue

        Model.ControlObjectFactoryNumber = View.FactoryNumberBox
        Set Model.Participants = ParticipantsFiller.GetParticipantsFromList()
        
        Set View = Nothing
        Transfer.Execute Model
    
    End Sub
        Private Function GetSubcontractValue() As String
                    
            Dim Names As List_String_
            Set Names = SubcontractListSelector _
                        .GetSelectedValues(View.SubcontractList)
                    
            Dim Result As String
            Result = SubcontractHandler.GetValue(Names)
                    
            GetSubcontractValue = Result
                    
        End Function
'---------------------------------------------
Private Sub ITransferDialogAction_HideDialog()
    OnDialogHide
End Sub

Private Sub ITransferDialogAction_InitiateTransfer()
    OnTransferInitiated
End Sub


