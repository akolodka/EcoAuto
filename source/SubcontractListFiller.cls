VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SubcontractListFiller"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder "eco_transferToWord.Core.Dialogs.Main Transfer.Presenter Delegate"
'@PredeclaredId
Option Explicit
Implements IPresenterDelegate

Private Type TSubcontractListFiller

    List As MSForms.ListBox
    Subcontracts As List_IFile_
    
End Type

Private this As TSubcontractListFiller

Public Function Create( _
       ByVal List As MSForms.ListBox, _
       ByVal Subcontracts As List_IFile_) As IPresenterDelegate

    With New SubcontractListFiller
    
        Set .List = List
        Set .Subcontracts = Subcontracts
        
        Set Create = .Self
    
    End With

End Function

Public Property Get Self() As IPresenterDelegate
    Set Self = Me
End Property

Public Property Get List() As MSForms.ListBox
    Set List = this.List
End Property
Public Property Set List(ByVal RHS As MSForms.ListBox)
    Set this.List = RHS
End Property

Public Property Get Subcontracts() As List_IFile_
    Set Subcontracts = this.Subcontracts
End Property
Public Property Set Subcontracts(ByVal RHS As List_IFile_)
    Set this.Subcontracts = RHS
End Property

Private Sub Execute()
    
    List.Clear
    
    Dim i As Long
    For i = 1 To Subcontracts.Count
    
        Process Subcontracts(i)

    Next
        
End Sub

Private Sub Process( _
        ByVal File As IFile)
        
    Dim Name As String
    Name = File.NameWithoutExtension
        
    If (IsSubcontactShouldBeShown(Name) = True) Then _
       List.AddItem Name

End Sub

Private Function IsSubcontactShouldBeShown( _
        ByVal NameMask As String) As Boolean
            
    IsSubcontactShouldBeShown = (NameMask Like "_*" = False)
            
End Function

'---------------------------------------------
Private Sub IPresenterDelegate_Execute()
    Execute
End Sub


