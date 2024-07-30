VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ExcelWorkbooksServiceFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("eco_transferToWord.Core.Excel")
Option Explicit
Implements IExcelWorkbooksServiceFactory

Private Function IExcelWorkbooksServiceFactory_Create( _
        ByVal Fso As IFileSystemObject) As IExcelWorkbooksService
    
    With New ExcelWorkbooksService
    
        Set .Fso = Fso
        Set IExcelWorkbooksServiceFactory_Create = .Self
        
    End With

End Function


