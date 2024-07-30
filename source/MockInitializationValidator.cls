VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MockInitializationValidator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder("Tests.Mocks.Initialization.Validator")
'@PredeclaredId
Option Explicit

Public Property Get CreateAsSuccess() As IInitializationValidator
    
    With New InitializationValidator
        
        .TemplatesFilesCount = 1
        .StaticValuesFilesCount = 1
        .SubcontractFilesCount = 1
        .ComparisonAuxillaryFilesCount = 1
        .ParticipantsCount = 1
        
        Set CreateAsSuccess = .Self
        
    End With
    
End Property

Public Property Get CreateAsWarning() As IInitializationValidator
    
    With New InitializationValidator
        
        .TemplatesFilesCount = 1
        .StaticValuesFilesCount = 1
        
        Set CreateAsWarning = .Self
        
    End With
    
End Property

Public Property Get CreateAsFailure() As IInitializationValidator
    
    With New InitializationValidator
        Set CreateAsFailure = .Self
    End With
    
End Property

