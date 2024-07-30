VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DecorationTableHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder "eco_transferToWord.Core.Comparison Results Decoration.Decoration Table Handler"
'@PredeclaredId
Option Explicit
Implements IDecorationTableHandler

Private Type TDecorationTableHandler
    Document As IEditableWordDocument
End Type

Private this As TDecorationTableHandler

Public Function Create( _
       ByVal Document As IEditableWordDocument) As IDecorationTableHandler
       
    With New DecorationTableHandler
    
        Set .Document = Document
        Set Create = .Self
    
    End With

End Function

Public Property Get Document() As IEditableWordDocument
    Set Document = this.Document
End Property
Public Property Set Document(ByVal RHS As IEditableWordDocument)
    Set this.Document = RHS
End Property

Public Property Get Self() As IDecorationTableHandler
    Set Self = Me
End Property

Private Function GetTable( _
        Optional ByVal Index As Long = vbEmpty) As IWordTable
    
    If (Document Is Nothing = True) Then _
        Exit Function
    
    If (Document.Range.Tables.Count = vbEmpty) Then _
        Exit Function

    If (Index > Document.Range.Tables.Count) Then _
        Exit Function
        
    Dim Result As IWordTable
    Set Result = Document.Range.LastTable
    
    If Index > vbEmpty Then _
        Set Result = Document.Range.Tables(Index)
    
    Set GetTable = Result
    
End Function
'---------------------------------------------
Private Function IDecorationTableHandler_GetTable(Optional ByVal Index As Long = vbEmpty) As IWordTable
    Set IDecorationTableHandler_GetTable = GetTable(Index)
End Function


