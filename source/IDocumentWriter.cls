VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IDocumentWriter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "eco_transferToWord.Core.Document IO.Writer"
'@Interface
Option Explicit

Public Enum InsertType
    
    ReplaceAll
    ReplaceFirstEntry
    ReplaceTwice
    
End Enum

Public Sub InsertIntoDocument( _
       ByVal KeyValuePair As IKeyValuePair, _
       Optional ByVal InsertOption As InsertType = InsertType.ReplaceAll)

End Sub



