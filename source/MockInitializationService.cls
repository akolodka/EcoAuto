VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MockInitializationService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder "Tests.Mocks.Initialization"
'@PredeclaredId
Option Explicit

Public Function Create() As IInitializationService
    
    Dim Preset As InitializationPresetDTO
    Set Preset = New InitializationPresetDTO
    
    Set Preset.Configuration = MockConfigurationService.Create()
    Set Preset.FileSystemObject = MockFileSystemObjectService.Create()
    
    Set Preset.ProgressPrimary = New MockProgressService
    
    Set Preset.WordApplicationFactory = New MockWordApplicationServiceFacto
    Set Preset.WorkbooksFactory = New MockExcelWorkbooksServiceFactor
    
    Set Create = InitializationService.Create(Preset)
        
End Function

Public Function CreateBattleSilent() As IInitializationService
    
    Dim Preset As InitializationPresetDTO
    Set Preset = New InitializationPresetDTO
    
    Set Preset.ProgressPrimary = New MockProgressService
    
    Set CreateBattleSilent = InitializationService.Create(Preset)
    
End Function


