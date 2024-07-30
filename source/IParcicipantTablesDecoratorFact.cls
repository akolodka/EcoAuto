VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IParcicipantTablesDecoratorFact"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "eco_transferToWord.Core.Comparison Results Decoration.Participant Tables Decoration"
'@Interface
Option Explicit

Public Function Create( _
       ByVal Strategy As IComparisonDecorationStrategy, _
       ByVal ParticipantTableTemplateFiles As List_IFile_, _
       ByVal ParticipantResults As List_IParticipantResult_) As IParticipantTablesDecorator
End Function


