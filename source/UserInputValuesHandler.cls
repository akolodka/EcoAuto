VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "UserInputValuesHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("eco_transferToWord.Core.Transfer Process.Values Handling.User Input")
Option Explicit
Implements IValuesHandler

Private Type TUserInputValuesHandler

    Writer As IDocumentWriter

    Subcontract As IKeyValuePair
    Respondent As IKeyValuePair
    
End Type

Private this As TUserInputValuesHandler

Public Property Get Self() As IValuesHandler
    Set Self = Me
End Property

Public Property Get Subcontract() As IKeyValuePair
    Set Subcontract = this.Subcontract
End Property
Public Property Set Subcontract(ByVal RHS As IKeyValuePair)
    Set this.Subcontract = RHS
End Property

Public Property Get Respondent() As IKeyValuePair
    Set Respondent = this.Respondent
End Property
Public Property Set Respondent(ByVal RHS As IKeyValuePair)
    Set this.Respondent = RHS
End Property

Public Property Get Writer() As IDocumentWriter
    Set Writer = this.Writer
End Property
Public Property Set Writer(ByVal RHS As IDocumentWriter)
    Set this.Writer = RHS
End Property

Public Property Get IsInitializedCorrect() As Boolean
    IsInitializedCorrect = (Subcontract.Value <> vbNullString) And (Respondent.Value <> vbNullString)
End Property

Public Sub ProcessValues()
    
    If (IsInitializedCorrect = False) Then _
       Exit Sub
    
    Process Subcontract
    Process Respondent
    
End Sub
    Private Sub Process( _
            ByVal Pair As IKeyValuePair)
            
        If (Pair Is Nothing = True) Then _
           Exit Sub
            
        Writer.InsertIntoDocument Pair
                        
    End Sub
'---------------------------------------------
Private Sub IValuesHandler_ProcessValues()
    ProcessValues
End Sub


