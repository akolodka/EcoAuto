VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ReferenceValueValidationDTO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "eco_transferToWord.Core.Reference Values Validator.Model.Reference Value"
Option Explicit
Implements IReferenceValueValidationModel

Private Type TReferenceValueValidationModel

    Value As String
    Uncertainty As String
    ParticipantNames As List_String_
    
End Type

Private this As TReferenceValueValidationModel

Public Property Get Self() As IReferenceValueValidationModel
    Set Self = Me
End Property
 
Public Property Get Value() As String
    Value = this.Value
End Property
Public Property Let Value(ByVal RHS As String)
    this.Value = RHS
End Property

Public Property Get Uncertainty() As String
    Uncertainty = this.Uncertainty
End Property
Public Property Let Uncertainty(ByVal RHS As String)
    this.Uncertainty = RHS
End Property

Public Property Get ParticipantNames() As List_String_
    Set ParticipantNames = this.ParticipantNames
End Property
Public Property Set ParticipantNames(ByVal RHS As List_String_)
    Set this.ParticipantNames = RHS
End Property
'---------------------------------------------
Private Property Get IReferenceValueValidationModel_Value() As String
    IReferenceValueValidationModel_Value = Value
End Property
Private Property Get IReferenceValueValidationModel_Uncertainty() As String
    IReferenceValueValidationModel_Uncertainty = Uncertainty
End Property
Private Property Get IReferenceValueValidationModel_ParticipantNames() As List_String_
    Set IReferenceValueValidationModel_ParticipantNames = ParticipantNames
End Property
