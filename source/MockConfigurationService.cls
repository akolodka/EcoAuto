VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MockConfigurationService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder "Tests.Mocks.Configuration"
'@PredeclaredId
Option Explicit

Public Function Create() As IConfiguration
    
    With New ConfigurationDTO
        
        .TemporaryFolderPath = "Temporary FolderPath"
        .ResultFolderPath = "Result FolderPath"
        
        .TemplatesFolderPath = "Templates FolderPath"
        .StaticDataFolderPath = "StaticData FolderPath"
        .SubcontractFolderPath = "Subcontracts FolderPath"
        .ComparisonAuxillaryFolderPath = "ComparisonAuxillaryFiles FolderPath"
        
        .WorksheetName = "WorkSheet Name"
        .MeasuredValuesRangesWorkbookPath = "MeasuredValuesRangesWorkbook Path"
        
        Set Create = .Self
        
    End With
    
End Function

