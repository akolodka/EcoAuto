VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IResultInterpretationHandlerFac"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("eco_transferToWord.Core.Transfer Process.Result Interpretation Handler")
'@Interface
Option Explicit

Public Function Create( _
       ByVal Document As IWordDocument, _
       ByVal Results As List_IParticipantResult_) As IResultInterpretationHandler

End Function


