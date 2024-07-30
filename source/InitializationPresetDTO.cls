VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "InitializationPresetDTO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("eco_transferToWord.Core.Initialization.Preset")
Option Explicit
Implements IInitializationPreset

Private Type TInitializationPresetDTO

    Configuration As IConfiguration
    FileSystemObject As IFileSystemObject
    ProgressPrimary As IProgressServicePrimary
    WordApplicationFactory As IWordApplicationServiceFactory
    WorkbooksFactory As IExcelWorkbooksServiceFactory
    
End Type

Private this As TInitializationPresetDTO

Public Property Get Configuration() As IConfiguration
    Set Configuration = this.Configuration
End Property
Public Property Set Configuration(ByVal RHS As IConfiguration)
    Set this.Configuration = RHS
End Property

Public Property Get FileSystemObject() As IFileSystemObject
    Set FileSystemObject = this.FileSystemObject
End Property
Public Property Set FileSystemObject(ByVal RHS As IFileSystemObject)
    Set this.FileSystemObject = RHS
End Property

Public Property Get ProgressPrimary() As IProgressServicePrimary
    Set ProgressPrimary = this.ProgressPrimary
End Property
Public Property Set ProgressPrimary(ByVal RHS As IProgressServicePrimary)
    Set this.ProgressPrimary = RHS
End Property

Public Property Get WordApplicationFactory() As IWordApplicationServiceFactory
    Set WordApplicationFactory = this.WordApplicationFactory
End Property
Public Property Set WordApplicationFactory(ByVal RHS As IWordApplicationServiceFactory)
    Set this.WordApplicationFactory = RHS
End Property

Public Property Get WorkbooksFactory() As IExcelWorkbooksServiceFactory
    Set WorkbooksFactory = this.WorkbooksFactory
End Property
Public Property Set WorkbooksFactory(ByVal RHS As IExcelWorkbooksServiceFactory)
    Set this.WorkbooksFactory = RHS
End Property
'---------------------------------------------
Private Property Get IInitializationPreset_Configuration() As IConfiguration
    Set IInitializationPreset_Configuration = Configuration
End Property
Private Property Get IInitializationPreset_FileSystemObject() As IFileSystemObject
    Set IInitializationPreset_FileSystemObject = FileSystemObject
End Property
Private Property Get IInitializationPreset_ProgressPrimary() As IProgressServicePrimary
    Set IInitializationPreset_ProgressPrimary = ProgressPrimary
End Property
Private Property Get IInitializationPreset_WordApplicationFactory() As IWordApplicationServiceFactory
    Set IInitializationPreset_WordApplicationFactory = WordApplicationFactory
End Property
Private Property Get IInitializationPreset_WorkbooksFactory() As IExcelWorkbooksServiceFactory
    Set IInitializationPreset_WorkbooksFactory = WorkbooksFactory
End Property
