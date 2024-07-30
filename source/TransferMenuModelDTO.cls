VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TransferMenuModelDTO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "eco_transferToWord.Core.Dialogs.Main Transfer.Model"
Option Explicit

Implements ITransferMenuModel
Implements IUserInputTransferModel
Implements IExcelTransferModel

Private Type TTransferMenuModelDTO

    Respondent As String
    Subcontract As String
    
    Participants As List_IParticipant_
    ControlObjectFactoryNumber As String
    
End Type

Private this As TTransferMenuModelDTO

Public Property Get Respondent() As String
    Respondent = this.Respondent
End Property
Public Property Let Respondent(ByVal Value As String)
    this.Respondent = Value
End Property

Public Property Get Subcontract() As String
    Subcontract = this.Subcontract
End Property
Public Property Let Subcontract(ByVal Value As String)
    this.Subcontract = Value
End Property

Public Property Get Participants() As List_IParticipant_
    Set Participants = this.Participants
End Property
Public Property Set Participants(ByVal RHS As List_IParticipant_)
    Set this.Participants = RHS
End Property

Public Property Get ControlObjectFactoryNumber() As String
    ControlObjectFactoryNumber = this.ControlObjectFactoryNumber
End Property
Public Property Let ControlObjectFactoryNumber(ByVal RHS As String)
    this.ControlObjectFactoryNumber = RHS
End Property
Public Property Get Self() As ITransferMenuModel
    Set Self = Me
End Property
'---------------------------------------------
Private Property Get IUserInputTransferModel_Respondent() As String
    IUserInputTransferModel_Respondent = Respondent
End Property
Private Property Get IUserInputTransferModel_Subcontract() As String
    IUserInputTransferModel_Subcontract = Subcontract
End Property
'---------------------------------------------
Private Property Get IExcelTransferModel_Participants() As List_IParticipant_
    Set IExcelTransferModel_Participants = Participants
End Property
Private Property Get IExcelTransferModel_ControlObjectFactoryNumber() As String
    IExcelTransferModel_ControlObjectFactoryNumber = ControlObjectFactoryNumber
End Property
