VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ComparisonBlockValidationDTO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "eco_transferToWord.Core.Reference Values Validator.Model.Comparison Block"
Option Explicit

Implements IComparisonBlockValidationModel
Implements IUniqueableComparisonBlockModel
Implements INotifiableComparisonBlockModel
Implements ISheetSelectableComparisonModel

Private Type TComparisonBlockValidationDTO

    IsReferenceValueUnuque As Boolean
    
    BlockIndex As Long
    ParticipantIndex As Long
    
    BlockValueDesctiption As String
    ParticipantName As String
    
End Type

Private this As TComparisonBlockValidationDTO

Public Property Get BlockIndex() As Long
    BlockIndex = this.BlockIndex
End Property
Public Property Let BlockIndex(ByVal RHS As Long)
    this.BlockIndex = RHS
End Property

Public Property Get ParticipantIndex() As Long
    ParticipantIndex = this.ParticipantIndex
End Property
Public Property Let ParticipantIndex(ByVal RHS As Long)
    this.ParticipantIndex = RHS
End Property

Public Property Get BlockValueDesctiption() As String
    BlockValueDesctiption = this.BlockValueDesctiption
End Property
Public Property Let BlockValueDesctiption(ByVal RHS As String)
    this.BlockValueDesctiption = RHS
End Property

Public Property Get ParticipantName() As String
    ParticipantName = this.ParticipantName
End Property
Public Property Let ParticipantName(ByVal RHS As String)
    this.ParticipantName = RHS
End Property

Public Property Get IsReferenceValueUnuque() As Boolean
    IsReferenceValueUnuque = this.IsReferenceValueUnuque
End Property
Public Property Let IsReferenceValueUnuque(ByVal RHS As Boolean)
    this.IsReferenceValueUnuque = RHS
End Property

Public Property Get Self() As IComparisonBlockValidationModel
    Set Self = Me
End Property
'---------------------------------------------
Private Property Get IUniqueableComparisonBlockModel_IsDataUnique() As Boolean
    IUniqueableComparisonBlockModel_IsDataUnique = IsReferenceValueUnuque
End Property
'---------------------------------------------
Private Property Get INotifiableComparisonBlockModel_BlockValueDesctiption() As String
    INotifiableComparisonBlockModel_BlockValueDesctiption = BlockValueDesctiption
End Property
Private Property Get INotifiableComparisonBlockModel_ParticipantName() As String
    INotifiableComparisonBlockModel_ParticipantName = ParticipantName
End Property
'---------------------------------------------
Private Property Get ISheetSelectableComparisonModel_BlockIndex() As Long
    ISheetSelectableComparisonModel_BlockIndex = BlockIndex
End Property
Private Property Get ISheetSelectableComparisonModel_ParticipantIndex() As Long
    ISheetSelectableComparisonModel_ParticipantIndex = ParticipantIndex
End Property
