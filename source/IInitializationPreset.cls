VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IInitializationPreset"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("eco_transferToWord.Core.Initialization.Preset")
'@Interface
Option Explicit

Public Property Get Configuration() As IConfiguration
End Property

Public Property Get FileSystemObject() As IFileSystemObject
End Property

Public Property Get ProgressPrimary() As IProgressServicePrimary
End Property

Public Property Get WordApplicationFactory() As IWordApplicationServiceFactory
End Property

Public Property Get WorkbooksFactory() As IExcelWorkbooksServiceFactory
End Property
