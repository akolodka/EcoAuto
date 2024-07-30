VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ValidationResultNotifier"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder "eco_transferToWord.Core.Reference Values Validator.Validation Result Notifier"
'@PredeclaredId
Option Explicit
Implements IValidationResultNotifier

Private Type TValidationResultNotifier

    Message As String
    ValidationSummary As List_IComparisonBlockValidMode_
    
End Type

Private this As TValidationResultNotifier

Public Function Create( _
       ByVal Summary As List_IComparisonBlockValidMode_) As IValidationResultNotifier
    
    With New ValidationResultNotifier
        
        Set .ValidationSummary = Summary
        .Initialize
    
        Set Create = .Self
        
    End With
    
End Function

Public Property Get Message() As String
    Message = this.Message
End Property
Public Property Let Message(ByVal RHS As String)
    this.Message = RHS
End Property

Public Property Get ValidationSummary() As List_IComparisonBlockValidMode_
    Set ValidationSummary = this.ValidationSummary
End Property
Public Property Set ValidationSummary(ByVal RHS As List_IComparisonBlockValidMode_)
    Set this.ValidationSummary = RHS
End Property

Public Property Get Self() As IValidationResultNotifier
    Set Self = Me
End Property

Public Sub Initialize()

    Dim Result As String
    Result = "Обнаружено расхождение данных приписанного значения:"

    Dim i As Long
    For i = 1 To ValidationSummary.Count
        
        Dim Model As IUniqueableComparisonBlockModel
        Set Model = ValidationSummary(i)
        
        If (Model.IsDataUnique = False) Then
                
            If Result <> vbNullString Then _
                Result = Result & vbNewLine & vbNewLine
            
            Dim CommentText As String
            CommentText = GetCommentText(Model)
            
            Result = Result & CommentText
            
        End If
        
    Next
    
    Result = Result & vbNewLine & vbNewLine & "Для переноса данных устраните несоответствие."
    Message = Result

End Sub
    Private Function GetCommentText( _
            ByVal Model As INotifiableComparisonBlockModel) As String
        
        Dim Prefix As String
        Prefix = "Блок: " & Model.BlockValueDesctiption & "." & vbNewLine
        
        Dim Body As String
        Body = "Участник: " & Model.ParticipantName & "."

        Dim Result As String
        Result = Prefix & Body
        
        GetCommentText = Result
        
    End Function
Public Sub SendMessage()
    MsgBox Message, vbInformation
End Sub
'---------------------------------------------
Private Sub IValidationResultNotifier_SendMessage()
    SendMessage
End Sub


