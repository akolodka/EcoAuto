VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IExcelValuesHandlerFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "eco_transferToWord.Core.Transfer Process.Values Handling.Excel"
'@Interface
Option Explicit

Public Function Create(ByVal DocumentWriter As IDocumentWriter, ByVal ParticipantILC As IComparisonDataParticipant) As IValuesHandler
End Function

