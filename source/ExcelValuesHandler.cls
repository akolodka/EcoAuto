VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ExcelValuesHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("eco_transferToWord.Core.Transfer Process.Values Handling.Excel")
Option Explicit
Implements IValuesHandler

Private Type TExcelValuesHandler

    Participant As IComparisonDataParticipant
    Writer As IDocumentWriter
    
    AuxillaryFilesHandlerFactory As IValuesHandlerFactory

End Type

Private this As TExcelValuesHandler

Public Property Get Self() As IValuesHandler
    Set Self = Me
End Property

Public Property Get Participant() As IComparisonDataParticipant
    Set Participant = this.Participant
End Property
Public Property Set Participant(ByVal RHS As IComparisonDataParticipant)
    Set this.Participant = RHS
End Property

Public Property Get Writer() As IDocumentWriter
    Set Writer = this.Writer
End Property
Public Property Set Writer(ByVal RHS As IDocumentWriter)
    Set this.Writer = RHS
End Property

Public Property Get AuxillaryFilesHandlerFactory() As IValuesHandlerFactory
    Set AuxillaryFilesHandlerFactory = this.AuxillaryFilesHandlerFactory
End Property
Public Property Set AuxillaryFilesHandlerFactory(ByVal RHS As IValuesHandlerFactory)
    Set this.AuxillaryFilesHandlerFactory = RHS
End Property

Public Property Get IsInitializedCorrect() As Boolean
    
    IsInitializedCorrect = (Writer Is Nothing = False) _
                           And (Participant Is Nothing = False) _
                           And (AuxillaryFilesHandlerFactory Is Nothing = False)

End Property

Public Sub ProcessValues()
    
    If (IsInitializedCorrect = False) Then _
        Exit Sub
    
    TransferAuxillaryFilesContent
    
    Process Participant.Number
    Process Participant.Organisation
    Process Participant.TypeP

End Sub
    Private Sub TransferAuxillaryFilesContent()
    
        Dim Handler As IValuesHandler
        Set Handler = AuxillaryFilesHandlerFactory.Create(Writer)
        
        Handler.ProcessValues
    
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

