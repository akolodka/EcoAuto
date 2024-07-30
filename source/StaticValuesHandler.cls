VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "StaticValuesHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("eco_transferToWord.Core.Transfer Process.Values Handling.Static")
Option Explicit
Implements IValuesHandler

Private Type TStaticValuesHandler
        
    KeyValuePairs As List_IKeyValuePair_
    Writer As IDocumentWriter

End Type

Private this As TStaticValuesHandler

Public Property Get Self() As IValuesHandler
    Set Self = Me
End Property

Public Property Get Writer() As IDocumentWriter
    Set Writer = this.Writer
End Property
Public Property Set Writer(ByVal RHS As IDocumentWriter)
    Set this.Writer = RHS
End Property

Public Property Get KeyValuePairs() As List_IKeyValuePair_
    Set KeyValuePairs = this.KeyValuePairs
End Property
Public Property Set KeyValuePairs(ByVal RHS As List_IKeyValuePair_)
    Set this.KeyValuePairs = RHS
End Property

Public Property Get IsInitializedCorrect() As Boolean
    
    If (KeyValuePairs Is Nothing = True) Or (Writer Is Nothing = True) Then _
        Exit Property
    
    IsInitializedCorrect = (KeyValuePairs.Count > vbEmpty)
    
End Property

Public Sub ProcessValues()
    
    If (IsInitializedCorrect = False) Then _
       Exit Sub
    
    Dim i As Long
    For i = 1 To KeyValuePairs.Count

        Process KeyValuePairs(i)

    Next

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

