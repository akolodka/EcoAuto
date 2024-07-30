VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IComparisonChartModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("eco_transferToWord.Core.Initialization.Comparison Mapping.Model.Participant Results.Contracts")
'@Interface
Option Explicit

Public Property Get ValueName() As String
End Property

Public Property Get ParticipantValue() As String
End Property

Public Property Get ParticipantUncertainty() As String
End Property

Public Property Get ReferenceValue() As String
End Property

Public Property Get ReferenceUncertainty() As String
End Property
