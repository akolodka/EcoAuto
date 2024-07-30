VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SubcontractListSelector"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder "eco_transferToWord.Core.Dialogs.Main Transfer.Presenter Delegate"
'@PredeclaredId
Option Explicit
Implements IPresenterDelegate

Private Type TSubcontractListSelector

    List As MSForms.ListBox
    Selected As List_String_
    
End Type

Private this As TSubcontractListSelector

Public Function Create( _
       ByVal List As MSForms.ListBox, _
       ByVal SelectedList As List_String_) As IPresenterDelegate

    With New SubcontractListSelector
    
        Set .List = List
        Set .Selected = SelectedList
        
        Set Create = .Self
    
    End With

End Function

Public Property Get List() As MSForms.ListBox
    Set List = this.List
End Property
Public Property Set List(ByVal RHS As MSForms.ListBox)
    Set this.List = RHS
End Property

Public Property Get Selected() As List_String_
    Set Selected = this.Selected
End Property
Public Property Set Selected(ByVal RHS As List_String_)
    Set this.Selected = RHS
End Property

Public Property Get Self() As IPresenterDelegate
    Set Self = Me
End Property

Private Sub Execute()

    Dim i As Long
    For i = 1 To Selected.Count
        
        ProcessSelection Selected(i)
        
    Next

End Sub

Private Function ProcessSelection( _
        ByVal Index As Long)
        
    If (IsSelectable(Index) = True) Then _
       List.Selected(Index) = True
        
End Function

Private Function IsSelectable( _
        ByVal Index As Long) As Boolean
            
    IsSelectable = (Index <= List.ListCount)

End Function

Public Function GetSelectedIndexes( _
       ByVal List As MSForms.ListBox) As List_String_
             
    Dim Result As List_String_
    Set Result = New List_String_
    
    Dim i As Long
    For i = LBound(List.List) To UBound(List.List)
        
        If (List.Selected(i) = True) Then _
           Result.Add i
        
    Next
    
    Set GetSelectedIndexes = Result
    
End Function

Public Function GetSelectedValues( _
       ByVal List As MSForms.ListBox) As List_String_
    
    Dim Result As List_String_
    Set Result = New List_String_
    
    Dim i As Long
    For i = LBound(List.List) To UBound(List.List)
        
        If (List.Selected(i) = True) Then _
           Result.Add List.List(i)
        
    Next
    
    Set GetSelectedValues = Result
    
End Function

'---------------------------------------------
Private Sub IPresenterDelegate_Execute()
    Execute
End Sub


