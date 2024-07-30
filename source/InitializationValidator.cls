VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "InitializationValidator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder("eco_transferToWord.Core.Initialization.Validator")
'@PredeclaredId
Option Explicit
Implements IInitializationValidator

Private Type TInitializationValidator
    
    TemplatesFilesCount As Long
    StaticValuesFilesCount As Long
    SubcontractFilesCount As Long
    ComparisonAuxillaryFilesCount As Long
    ParticipantsCount As Long

End Type

Private this As TInitializationValidator

Public Property Get Create( _
       ByVal Initial As IValidatorInitialization) As IInitializationValidator
    
    With New InitializationValidator
        
        .TemplatesFilesCount = Initial.TemplateFilesCount
        .StaticValuesFilesCount = Initial.StaticValueFilesCount
        .SubcontractFilesCount = Initial.SubctontractFilesCount
        .ComparisonAuxillaryFilesCount = Initial.ComparisonAuxillaryFilesCount
        .ParticipantsCount = Initial.ParticipantsCount
        
        Set Create = .Self
        
    End With
    
End Property

Public Property Get TemplatesFilesCount() As Long
    TemplatesFilesCount = this.TemplatesFilesCount
End Property
Public Property Let TemplatesFilesCount(ByVal RHS As Long)
    this.TemplatesFilesCount = RHS
End Property

Public Property Get StaticValuesFilesCount() As Long
    StaticValuesFilesCount = this.StaticValuesFilesCount
End Property
Public Property Let StaticValuesFilesCount(ByVal RHS As Long)
    this.StaticValuesFilesCount = RHS
End Property

Public Property Get SubcontractFilesCount() As Long
    SubcontractFilesCount = this.SubcontractFilesCount
End Property
Public Property Let SubcontractFilesCount(ByVal RHS As Long)
    this.SubcontractFilesCount = RHS
End Property

Public Property Get ComparisonAuxillaryFilesCount() As Long
    ComparisonAuxillaryFilesCount = this.ComparisonAuxillaryFilesCount
End Property
Public Property Let ComparisonAuxillaryFilesCount(ByVal RHS As Long)
    this.ComparisonAuxillaryFilesCount = RHS
End Property

Public Property Get ParticipantsCount() As Long
    ParticipantsCount = this.ParticipantsCount
End Property
Public Property Let ParticipantsCount(ByVal RHS As Long)
    this.ParticipantsCount = RHS
End Property

Public Property Get Self() As IInitializationValidator
    Set Self = Me
End Property

Public Function GetValidationStatus() As IValidationStatus

    Set GetValidationStatus = New FailureNoTemplatesStatus
    
    If TemplatesFilesCount = vbEmpty Then _
       Exit Function
    
    Set GetValidationStatus = New FailureNoWorksheetStatus
    
    If ParticipantsCount = vbEmpty Then _
       Exit Function
    
    Set GetValidationStatus = New PartialReadyStatus
    
    If StaticValuesFilesCount = vbEmpty Then _
       Exit Function
    
    If SubcontractFilesCount = vbEmpty Then _
       Exit Function
    
    If ComparisonAuxillaryFilesCount = vbEmpty Then _
       Exit Function
       
    Set GetValidationStatus = New CompletelyReadyStatus
    
End Function

'---------------------------------------------
Private Function IInitializationValidator_GetValidationStatus() As IValidationStatus
    Set IInitializationValidator_GetValidationStatus = GetValidationStatus
End Function


