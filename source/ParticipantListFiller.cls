VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ParticipantListFiller"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder "eco_transferToWord.Core.Dialogs.Main Transfer.Presenter Delegate"
'@PredeclaredId
Option Explicit

Implements IParticipantListFiller
Implements IPresenterDelegate

Const IndexDelitimer = "                         #!"

Private Type TParticipantListFiller

    List As MSForms.ListBox
    Participants As List_IParticipant_
    
    FilterKey As String
    
End Type

Private this As TParticipantListFiller

Public Function Create( _
       ByVal List As MSForms.ListBox, _
       ByVal Participants As List_IParticipant_) As IParticipantListFiller

    With New ParticipantListFiller
        
        Set .List = List
        Set .Participants = Participants

        Set Create = .Self
    
    End With

End Function

Public Property Get List() As MSForms.ListBox
    Set List = this.List
End Property
Public Property Set List(ByVal RHS As MSForms.ListBox)
    Set this.List = RHS
End Property

Public Property Get Participants() As List_IParticipant_
    Set Participants = this.Participants
End Property
Public Property Set Participants(ByVal RHS As List_IParticipant_)
    Set this.Participants = RHS
End Property

Public Property Get FilterKey() As String
    FilterKey = this.FilterKey
End Property
Public Property Let FilterKey(ByVal RHS As String)
    this.FilterKey = RHS
End Property

Public Property Get Self() As ParticipantListFiller
    Set Self = Me
End Property

Private Sub Execute()
    
    List.Clear
    
    Dim i As Long
    For i = 1 To Participants.Count
    
        If (Predicate(Participants(i)) = True) Then _
           Process Participants(i), i
        
    Next
    
End Sub
    Private Function Predicate( _
            ByVal Participant As ITransferMenuParticipant) As Boolean
        
        Dim NumericPredicate As Boolean
        NumericPredicate = GetNumericPredicate(Participant.Tour)
        
        Dim StringPredicate As Boolean
        StringPredicate = GetStringPredicate(Participant.Organisation)
        
        Predicate = NumericPredicate Or StringPredicate
        
    End Function
        Private Function GetNumericPredicate( _
                ByVal TourNumber As String) As Boolean
                
            If (IsNumeric(FilterKey) = False) Then _
               Exit Function
                
            Dim Predicate As Boolean
            Predicate = (TourNumber = FilterKey) Or _
                        ("0" & TourNumber = FilterKey) Or _
                        (Len(TourNumber) = 1) And (FilterKey = "0")
                    
            GetNumericPredicate = Predicate
                
        End Function
        Private Function GetStringPredicate( _
                ByVal Organisation As String) As Boolean
                
            If (IsNumeric(FilterKey) = True) Then _
               Exit Function

            Dim ReversedKey As String
            ReversedKey = GetLayoutReversedKey(FilterKey)
                
            Dim Predicate As Boolean
            Predicate = (FilterKey = vbNullString) Or _
                        (InStr(LCase(Organisation), LCase(FilterKey)) > vbEmpty) Or _
                        (InStr(LCase(Organisation), LCase(ReversedKey)) > vbEmpty)

            GetStringPredicate = Predicate
                
        End Function
            Private Function GetLayoutReversedKey( _
                    ByVal FilterKey As String) As String
                                        
                Dim Reverser As ILayoutReverser
                Set Reverser = LayoutReverser.Create(FilterKey)
                
                Dim Result As String
                Result = Reverser.GetReversedString
                
                GetLayoutReversedKey = Result
                                        
            End Function
            
Private Sub Process( _
        ByVal Participant As ITransferMenuParticipant, _
        ByVal Index As Long)
        
    Dim Result As String
    Result = Format$(Participant.Tour, "0#") & "    " & _
             Participant.Organisation & IndexDelitimer & Index

    List.AddItem Result
        
End Sub

Private Function GetParticipants() As List_IParticipant_
    
    Dim Result As List_IParticipant_
    Set Result = GetSelectedParticipants
    
    If Result.Count = vbEmpty Then _
       Set Result = GetAllShowedParticipants
    
    Set GetParticipants = Result
    
End Function
    Private Function GetSelectedParticipants() As List_IParticipant_
            
        Dim Result As List_IParticipant_
        Set Result = New List_IParticipant_
            
        Dim i As Long
        For i = LBound(List.List) To UBound(List.List)
                
            If (List.Selected(i) = True) Then _
               Result.Add GetParticipantFromList(List.List(i))
                
        Next
            
        Set GetSelectedParticipants = Result
            
    End Function
        Private Function GetParticipantFromList( _
                ByVal Selected As String) As IComparisonDataParticipant
                    
            Dim temp() As String
            temp = Split(Selected, IndexDelitimer)
                    
            Dim Index As Long
            Index = CLng(temp(UBound(temp)))
                    
            If Index > Participants.Count Then _
               Exit Function
                    
            Dim Result As IComparisonDataParticipant
            Set Result = Participants(Index)
                    
            Set GetParticipantFromList = Result
            
        End Function

Private Function GetAllShowedParticipants() As List_IParticipant_
        
    Dim Result As List_IParticipant_
    Set Result = New List_IParticipant_
        
    Dim i As Long
    For i = LBound(List.List) To UBound(List.List)
        Result.Add GetParticipantFromList(List.List(i))
    Next
        
    Set GetAllShowedParticipants = Result
        
End Function
'---------------------------------------------
Private Function IParticipantListFiller_Create( _
        ByVal FilterKey As String) As IPresenterDelegate
    
    With New ParticipantListFiller
        
        Set .List = List
        Set .Participants = Participants
        
        .FilterKey = FilterKey

        Set IParticipantListFiller_Create = .Self
    
    End With
    
End Function
Private Function IParticipantListFiller_GetParticipantsFromList() As List_IParticipant_
    Set IParticipantListFiller_GetParticipantsFromList = GetParticipants
End Function
'---------------------------------------------
Private Sub IPresenterDelegate_Execute()
    Execute
End Sub


