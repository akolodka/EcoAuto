VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MockTransferMenuPresenter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder("Tests.Mocks.Transfer Menu")
'@PredeclaredId
Option Explicit
Implements ITransferDialogAction

Private Type TMockTransferMenuPresenter

    Transfer As ITransferService
    Participants As List_IParticipant_
    
End Type

Private this As TMockTransferMenuPresenter

Public Function Create( _
       ByVal Initial As ITransferDialogInitialization) As ITransferDialogAction

    With New MockTransferMenuPresenter
        
        Set .Transfer = Initial.TransferService
        Set .Participants = Initial.Participants

        Set Create = .Self
        
    End With
    
End Function

Public Property Get Transfer() As ITransferService
    Set Transfer = this.Transfer
End Property

Public Property Set Transfer(ByVal RHS As ITransferService)
    Set this.Transfer = RHS
End Property

Public Property Get Participants() As List_IParticipant_
    Set Participants = this.Participants
End Property

Public Property Set Participants(ByVal RHS As List_IParticipant_)
    Set this.Participants = RHS
End Property

Public Property Get Self() As ITransferDialogAction
    Set Self = Me
End Property

Private Sub ITransferDialogAction_HideDialog()
End Sub

Private Sub ITransferDialogAction_InitiateTransfer()
    
    Dim workList As List_IParticipant_
    Set workList = New List_IParticipant_
    
    workList.Add Participants(1)
    workList.Add Participants(2)

    Dim Model As TransferMenuModelDTO
    Set Model = New TransferMenuModelDTO
    
    Model.Respondent = "Пидоров М.Д."
    Model.Subcontract = "Здесь будут сведения о субподряде"
    
    Model.ControlObjectFactoryNumber = "ОК № 3"
    Set Model.Participants = workList

    Transfer.Execute Model
    
End Sub


