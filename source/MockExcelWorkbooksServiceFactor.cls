VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MockExcelWorkbooksServiceFactor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Tests.Mocks.Excel")
Option Explicit
Implements IExcelWorkbooksServiceFactory

Private Function IExcelWorkbooksServiceFactory_Create( _
        ByVal Fso As IFileSystemObject) As IExcelWorkbooksService

    With New MockExcelWorkbooksService
        Set IExcelWorkbooksServiceFactory_Create = .Self
    End With

End Function


