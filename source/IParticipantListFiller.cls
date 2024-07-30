VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IParticipantListFiller"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "eco_transferToWord.Core.Dialogs.Main Transfer.Presenter Delegate.Contracts"
'@Interface
Option Explicit

Public Function Create(ByVal FilterKey As String) As IPresenterDelegate
End Function

Public Function GetParticipantsFromList() As List_IParticipant_
End Function

