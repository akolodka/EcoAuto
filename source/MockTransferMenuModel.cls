VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MockTransferMenuModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder("Tests.Mocks.Transfer Menu")
'@PredeclaredId
Option Explicit

Public Function Create( _
       ByVal Initial As ITransferDialogInitialization) As ITransferMenuModel
    
    Dim workList As List_IParticipant_
    Set workList = New List_IParticipant_
    
    workList.Add Initial.Participants(1)
'    workList.Add Initial.Participants(2)
    workList.Add Initial.Participants(3)
    
    With New TransferMenuModelDTO
    
        .ControlObjectFactoryNumber = "Œ  π 2"
        Set .Participants = workList
        
        Set Create = .Self
    
    End With

End Function


