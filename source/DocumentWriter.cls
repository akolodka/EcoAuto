VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DocumentWriter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("eco_transferToWord.Core.Document IO.Writer")
Option Explicit
Implements IDocumentWriter

Private Type TDocumentWriter

    Progress As IProgressService
    StrategiesIO As List_ITransferValueIOStrategy_
    
    DocumentName As String
    Finder As IDocumentRangeFinder
    
End Type

Private this As TDocumentWriter

Public Property Get Self() As IDocumentWriter
    Set Self = Me
End Property

Public Property Get Progress() As IProgressService
    Set Progress = this.Progress
End Property
Public Property Set Progress(ByVal RHS As IProgressService)
    Set this.Progress = RHS
End Property

Public Property Get StrategiesIO() As List_ITransferValueIOStrategy_
    Set StrategiesIO = this.StrategiesIO
End Property
Public Property Set StrategiesIO(ByVal RHS As List_ITransferValueIOStrategy_)
    Set this.StrategiesIO = RHS
End Property

Public Property Get DocumentName() As String
    DocumentName = this.DocumentName
End Property
Public Property Let DocumentName(ByVal RHS As String)
    this.DocumentName = RHS
End Property

Public Property Get Finder() As IDocumentRangeFinder
    Set Finder = this.Finder
End Property
Public Property Set Finder(ByVal RHS As IDocumentRangeFinder)
    Set this.Finder = RHS
End Property

Private Sub InsertIntoDocument( _
        ByVal KeyValuePair As IKeyValuePair, _
        Optional ByVal InsertOption As InsertType = InsertType.ReplaceAll)

    If (InsertOption = InsertType.ReplaceAll) Then _
        Progress.Show "Заполнение шаблона «" & DocumentName & "»", "Ключ: " & KeyValuePair.Key
        
    Dim Writer As ITransferValueWriter
    Set Writer = GetWriteStrategy(KeyValuePair.Value)
    
    If (Writer Is Nothing = True) Then _
       Exit Sub
        
    Dim findedRange As IWordRange
    Set findedRange = Finder.Find(KeyValuePair.Key)
    
    If (findedRange Is Nothing = True) Then _
       Exit Sub
    
    Dim Destination As IWordRange
    Dim ExecutionCount As Long
    
    Dim TargetCount As Long
    TargetCount = GetTargetCount(InsertOption)
    
    Do
        ExecutionCount = ExecutionCount + 1
        Set Destination = Finder.Find(KeyValuePair.Key)

        If (Destination Is Nothing) Then _
           Exit Sub

        Writer.InsertIntoDocument KeyValuePair.Value, Destination
        
    Loop While (ExecutionCount < TargetCount)
    
End Sub
    Private Function GetTargetCount( _
            ByVal InsertOption As InsertType) As Long
        
        Dim Result As Long
        Result = 99999
        
        If (InsertOption = InsertType.ReplaceFirstEntry) Then _
            Result = 1
        
        If (InsertOption = InsertType.ReplaceTwice) Then _
            Result = 2
        
        GetTargetCount = Result
        
    End Function
    

    
    Private Function GetWriteStrategy( _
            ByVal Value As ITransferValue) As ITransferValueWriter
            
        Dim i As Long
        For i = 1 To StrategiesIO.Count
    
            If (StrategiesIO(i).IsMatchesForWrite(Value) = True) Then
        
                Set GetWriteStrategy = StrategiesIO(i)
                Exit Function
        
            End If
        
        Next
            
    End Function
'---------------------------------------------
Private Sub IDocumentWriter_InsertIntoDocument( _
        ByVal KeyValuePair As IKeyValuePair, _
        Optional ByVal InsertOption As InsertType = InsertType.ReplaceAll)
    
    InsertIntoDocument KeyValuePair, InsertOption
    
End Sub
