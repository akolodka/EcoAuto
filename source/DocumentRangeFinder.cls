VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DocumentRangeFinder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder "eco_transferToWord.Core.Document IO.Range Finder"
'@PredeclaredId
Option Explicit
Implements IDocumentRangeFinder

Private Type TDocumentRangeFinder
    Document As IEditableWordDocument
End Type

Private this As TDocumentRangeFinder

Public Function Create( _
       ByVal Document As IEditableWordDocument) As IDocumentRangeFinder
    
    With New DocumentRangeFinder
        'Важно передавать именно Document, потому что при поиске range происходит его замещение
        Set .Document = Document
        Set Create = .Self
        
    End With
      
End Function

Public Property Get Self() As IDocumentRangeFinder
    Set Self = Me
End Property

Public Property Get Document() As IEditableWordDocument
    Set Document = this.Document
End Property
Public Property Set Document(ByVal RHS As IEditableWordDocument)
    Set this.Document = RHS
End Property

'---------------------------------------------
Private Function IDocumentRangeFinder_Find( _
        ByVal Key As ITransferKey) As IWordRange
    
    If (Document Is Nothing = True) Then _
       Exit Function

    If (Key = vbNullString) Then _
       Exit Function
    
    Dim WorkRange As IFindableWordRange
    Set WorkRange = Document.Range
    
    Dim IsFinded As Boolean
    IsFinded = WorkRange.FindExecute(Key)

    If (IsFinded = True) Then _
       Set IDocumentRangeFinder_Find = WorkRange
            
End Function


