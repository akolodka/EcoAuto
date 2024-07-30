VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "List_IWordTable_"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "eco_transferToWord.Core.Strong Typed Lists"
Option Explicit

Private Type TList_IWordTable_
    List As Collection
End Type

Private this As TList_IWordTable_

Public Sub Add( _
       ByVal Value As IWordTable)
    
    If (List Is Nothing = True) Then _
       Set List = New Collection
    
    If (Value Is Nothing = True) Then _
       Exit Sub
    
    List.Add Value

End Sub

Private Property Get List() As Collection
    Set List = this.List
End Property
Private Property Set List(ByVal RHS As Collection)
    Set this.List = RHS
End Property

Public Property Get Count() As Long

    If (List Is Nothing = True) Then _
       Set List = New Collection
    
    Count = List.Count
    
End Property

'@DefaultMember
Public Property Get Item( _
       ByVal Index As Long) As IWordTable
Attribute Item.VB_UserMemId = 0
    
    If Index > Count Then _
       Exit Property
    
    Set Item = List(Index)
    
End Property