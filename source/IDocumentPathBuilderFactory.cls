VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IDocumentPathBuilderFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("eco_transferToWord.Core.Word.Path Builder.Contracts")
'@Interface
Option Explicit

Public Function Create(ByVal DocumentName As String, ByVal DocumentSuffix As String) As IDocumentPathBuilder
End Function

