VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MeasuredValuesHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "eco_transferToWord.Core.Preparation Document Templates.Values Ranges Filler.Measured Values Range Handler"
Option Explicit
Implements IMeasuredValuesHandler

Const DefaultValue As String = "< ? >"

Private Type TMeasuredValuesHandler

    Fso As IFileSystemObject
    WorkbookPath As String
    
    ControlObjectFactoryNumber As String
    ValuesRanges As List_IKeyValuePair_
    
End Type

Private this As TMeasuredValuesHandler

Public Property Get Self() As IMeasuredValuesHandler
    Set Self = Me
End Property

Public Property Get Fso() As IFileSystemObject
    Set Fso = this.Fso
End Property
Public Property Set Fso(ByVal RHS As IFileSystemObject)
    Set this.Fso = RHS
End Property

Public Property Get WorkbookPath() As String
    WorkbookPath = this.WorkbookPath
End Property
Public Property Let WorkbookPath(ByVal RHS As String)
    this.WorkbookPath = RHS
End Property

Public Property Get ControlObjectFactoryNumber() As String
    ControlObjectFactoryNumber = this.ControlObjectFactoryNumber
End Property
Public Property Let ControlObjectFactoryNumber(ByVal RHS As String)
    this.ControlObjectFactoryNumber = RHS
End Property

Public Property Get ValuesRanges() As List_IKeyValuePair_
    Set ValuesRanges = this.ValuesRanges
End Property
Public Property Set ValuesRanges(ByVal RHS As List_IKeyValuePair_)
    Set this.ValuesRanges = RHS
End Property

Public Sub Initialize()
       
    Dim TargetWorkbook As Excel.Workbook
    Set TargetWorkbook = GetTargetWorkbook()
    
    If (TargetWorkbook Is Nothing = True) Then _
        Set ValuesRanges = New List_IKeyValuePair_: _
        Exit Sub
          
    Dim wsData As Worksheet
    Set wsData = GetWorksheet(TargetWorkbook)
    
    Dim TableValuesRanges As ListObject
    Set TableValuesRanges = GetWorkTable(wsData)
            
    Set ValuesRanges = GetValuesRanges(TableValuesRanges)
    TargetWorkbook.Close False
    
End Sub
    Private Function GetTargetWorkbook() As Workbook
        
        If (Fso Is Nothing = True) Then _
           Exit Function
                   
        If (Fso.FileExists(WorkbookPath) = False) Then _
           Exit Function
           
        Set GetTargetWorkbook = GetObject(WorkbookPath)
        
    End Function
    Private Function GetWorksheet( _
            ByVal TargetWorkbook As Workbook) As Worksheet
            
        If (TargetWorkbook Is Nothing = True) Then _
            Exit Function
        
        If (TargetWorkbook.Worksheets.Count = vbEmpty) Then _
           Exit Function
    
        Dim Result As Worksheet
        Set Result = TargetWorkbook.Worksheets(1)
            
        Set GetWorksheet = Result
                
    End Function
    Private Function GetWorkTable( _
            ByVal Ws As Worksheet) As ListObject
        
        If (Ws Is Nothing = True) Then _
            Exit Function
        
        If (Ws.ListObjects.Count = vbEmpty) Then _
           Exit Function
            
        Dim Result As ListObject
        Set Result = Ws.ListObjects(1)
            
        Set GetWorkTable = Result
            
    End Function
    Private Function GetValuesRanges( _
            ByVal TableValuesRanges As ListObject) As List_IKeyValuePair_
        
        If (TableValuesRanges Is Nothing = True) Then _
            Exit Function
        
        Dim Result As List_IKeyValuePair_
        Set Result = New List_IKeyValuePair_
        
        Dim Ranges As Excel.Range
        Set Ranges = GetValuesRangesTableData(TableValuesRanges)
        
        Dim Names As Excel.Range
        Set Names = TableValuesRanges.ListColumns(1).Range
        
        Set Result = GetPairList(Names, Ranges)
        Set GetValuesRanges = Result
        
    End Function
        Private Function GetValuesRangesTableData( _
                ByVal TableValuesRanges As ListObject) As Excel.Range
            
            Dim FactoryNumbers As Excel.Range
            Set FactoryNumbers = TableValuesRanges.HeaderRowRange
            
            Dim IsFinded As Boolean
            Dim TargetColumn As Long
            
            Dim i As Long
            For i = LBound(FactoryNumbers.Value2, 2) To UBound(FactoryNumbers.Value2, 2)
                
                If (FactoryNumbers.Value2(1, i) = ControlObjectFactoryNumber) Then
                
                    IsFinded = True
                    TargetColumn = i
                    
                    Exit For
                    
                End If
                
            Next
    
            If (IsFinded = True) Then _
                Set GetValuesRangesTableData = TableValuesRanges.ListColumns(TargetColumn).Range
            
        End Function
        Private Function GetPairList( _
                ByVal Names As Excel.Range, _
                ByVal Ranges As Excel.Range) As List_IKeyValuePair_
                
            Dim Result As List_IKeyValuePair_
            Set Result = New List_IKeyValuePair_
            
            Set GetPairList = Result
            
            If (Names Is Nothing = True) Or _
               (Ranges Is Nothing = True) Then _
               Exit Function
    
            Dim i As Long
            For i = 1 To Names.Rows.Count
                
                Dim Pair As IKeyValuePair
                Set Pair = KeyValuePair.Create(Names.Value2(i, 1), Ranges.Value2(i, 1))
                
                Result.Add Pair
                    
            Next
            
            Set GetPairList = Result
                
        End Function
Public Function GetMeasuredValueRange( _
       ByVal Key As String) As String
    
    GetMeasuredValueRange = DefaultValue
    
    Dim workKey As ITransferKey
    Set workKey = TransferKey.Create(Key)
    
    Dim i As Long
    For i = 1 To ValuesRanges.Count
        
        If (ValuesRanges(i).Key = workKey) Then
        
            GetMeasuredValueRange = ValuesRanges(i).Value
            Exit Function
            
        End If
        
    Next
    
End Function
'---------------------------------------------
Private Function IMeasuredValuesHandler_GetMeasuredValueRange(ByVal Key As String) As String
    IMeasuredValuesHandler_GetMeasuredValueRange = GetMeasuredValueRange(Key)
End Function


