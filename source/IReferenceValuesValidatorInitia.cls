VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IReferenceValuesValidatorInitia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("eco_transferToWord.Core.Initialization.Contracts")
'@Interface
Option Explicit

Public Property Get Participants() As List_IParticipant_
End Property

Public Property Get SourceWorksheet() As IExcelWorksheet
End Property


