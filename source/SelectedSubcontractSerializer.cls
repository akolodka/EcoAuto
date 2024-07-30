VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SelectedSubcontractSerializer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "eco_transferToWord.Core.Dialogs.Main Transfer.Cache Handler.Serializer"
Option Explicit

Const Separator As String = ","

Public Function ToStringFromList( _
       ByVal Values As List_String_) As String
    
    If (Values Is Nothing = True) Then _
       Exit Function
    
    If (Values.Count = vbEmpty) Then _
       Exit Function
    
    Dim Result As String
    
    Dim i As Long
    For i = 1 To Values.Count
    
        If i > 1 Then _
           Result = Result & Separator
        
        Result = Result & Values(i)
    
    Next
    
    ToStringFromList = Result
    
End Function

Public Function DeserializeIntoList( _
       ByVal Value As String) As List_String_
    
    Dim Result As List_String_
    Set Result = New List_String_
    
    Set DeserializeIntoList = Result
     
    If Value = vbNullString Then _
       Exit Function
    
    Dim Values() As String
    Values = Split(Value, Separator)

    Dim i As Long
    For i = LBound(Values) To UBound(Values)
    
        Result.Add CLng(Values(i))
        
    Next
    
    Set DeserializeIntoList = Result
    
End Function


