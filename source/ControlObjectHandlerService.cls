VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ControlObjectHandlerService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder("eco_transferToWord.Core.Initialization.Control Object")
'@PredeclaredId
Option Explicit
Implements IControlObjectHandler

Private Type TControlObjectHandler

    Fso As IFileSystemObject
    WorkbookPath As String
    
    ControlObjectsNumbers As List_String_
    
End Type

Private this As TControlObjectHandler

Public Function Create( _
       ByVal Fso As IFileSystemObject, _
       ByVal WorkbookPath As String) As IControlObjectHandler
    
    With New ControlObjectHandlerService
        
        Set .Fso = Fso
        .WorkbookPath = WorkbookPath
        
        .Initialize
        Set Create = .Self
    
    End With

End Function

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

Public Property Get ControlObjectsNumbers() As List_String_
    Set ControlObjectsNumbers = this.ControlObjectsNumbers
End Property
Public Property Set ControlObjectsNumbers(ByVal RHS As List_String_)
    Set this.ControlObjectsNumbers = RHS
End Property

Public Property Get Self() As IControlObjectHandler
    Set Self = Me
End Property

Public Sub Initialize()
       
    Dim TargetWorkbook As Workbook
    Set TargetWorkbook = GetTargetWorkbook()
    
    If (TargetWorkbook Is Nothing = True) Then _
        Set ControlObjectsNumbers = New List_String_: _
        Exit Sub
          
    Dim wsData As Worksheet
    Set wsData = GetWorksheet(TargetWorkbook)
    
    Dim TableValuesRanges As ListObject
    Set TableValuesRanges = GetWorkTable(wsData)
            
    Set ControlObjectsNumbers = GetControlObjectsNumbers(TableValuesRanges)
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
    Private Function GetControlObjectsNumbers( _
            ByVal TableValuesRanges As ListObject) As List_String_

        If (TableValuesRanges Is Nothing = True) Then _
            Exit Function
            
        Dim Result As List_String_
        Set Result = New List_String_
                
        Dim FactoryNumbers As Excel.Range
        Set FactoryNumbers = TableValuesRanges.HeaderRowRange
        
        Dim i As Long
        For i = LBound(FactoryNumbers.Value2, 2) To UBound(FactoryNumbers.Value2, 2)
        
            If i > LBound(FactoryNumbers.Value2, 2) Then _
                Result.Add FactoryNumbers.Value2(1, i)
            
        Next

        Set GetControlObjectsNumbers = Result
        
    End Function
'---------------------------------------------
Private Property Get IControlObjectHandler_FactoryNumbers() As List_String_
    Set IControlObjectHandler_FactoryNumbers = ControlObjectsNumbers
End Property


