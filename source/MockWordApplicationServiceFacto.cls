VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MockWordApplicationServiceFacto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Tests.Mocks.Word.Application.Factory")
Option Explicit
Implements IWordApplicationServiceFactory

Private Function IWordApplicationServiceFactory_Create( _
    ByVal Progress As IProgressService) As IWordApplicationService

    With New MockWordApplicationService
        
        Set .Progress = Progress
        Set IWordApplicationServiceFactory_Create = .Self
        
    End With

End Function

