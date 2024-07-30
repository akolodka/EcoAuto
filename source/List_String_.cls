VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "List_String_"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "eco_transferToWord.Core.Strong Typed Lists"
Option Explicit

Private Type TList_String_
    List As Collection
End Type

Private this As TList_String_

Public Sub Add( _
       ByVal Value As String)
    
    If (List Is Nothing = True) Then _
       Set List = New Collection
    
    If Value = vbNullString Then _
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
       ByVal Index As Long) As String
Attribute Item.VB_UserMemId = 0
    
    If Index > Count Then _
       Exit Property
    
    Item = List(Index)
    
End Property

Public Function EqualsTo( _
       ByVal Reference As List_String_) As Boolean
    
    If List.Count <> Reference.Count Then _
       Exit Function
        
    Dim i As Long
    For i = 1 To List.Count
    
        If List(i) <> Reference(i) Then _
           Exit Function
    
    Next
    
    EqualsTo = True

End Function
