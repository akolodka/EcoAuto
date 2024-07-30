VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MockTransferService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder("Tests.Mocks")
'@PredeclaredId
Option Explicit
Implements ITransferService

Private Type TMockTransferService
    Model As IExcelTransferModel
End Type

Private this As TMockTransferService

Public Function Create() As ITransferService
    
    With New MockTransferService
        Set Create = .Self
    End With

End Function

Public Property Get Model() As IExcelTransferModel
    Set Model = this.Model
End Property

Public Property Set Model(ByVal RHS As IExcelTransferModel)
    Set this.Model = RHS
End Property

Public Property Get Self() As ITransferService
    Set Self = Me
End Property

Private Sub ITransferService_Execute( _
        ByVal ViewModel As ITransferMenuModel)
    
    Set Model = ViewModel
    
    Dim i As Long
    For i = 1 To Model.Participants.Count
        
        Dim Participant As IComparisonDataParticipant
        Set Participant = Model.Participants(i)
        
Debug.Print
Debug.Print i; Participant.Number
Debug.Print " "; Participant.Organisation
Debug.Print " "; Participant.Tour
    
    Next
    
End Sub


