VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ExcelWorkbooksService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("eco_transferToWord.Core.Excel")
Option Explicit

Implements IExcelWorkbooksService
Implements IDisposableExcelWorbooksService

Private Type TExcelWorkbooksService
    Fso As IFileSystemObject
End Type

Private this As TExcelWorkbooksService

Public Property Get Fso() As IFileSystemObject
    Set Fso = this.Fso
End Property
Public Property Set Fso(ByVal RHS As IFileSystemObject)
    Set this.Fso = RHS
End Property

Public Property Get Self() As IExcelWorkbooksService
    Set Self = Me
End Property

Public Property Get ActiveBook() As IExcelWorkbook
    
    Dim Result As IExcelWorkbook
    Set Result = ExcelWorkbook.Create(ActiveWorkbook)
    
    Set ActiveBook = Result
    
End Property

Public Function GetFrom( _
       ByVal WorkbookPath As String) As IExcelWorkbook
    
    If (Fso.FileExists(WorkbookPath) = False) Then _
        Exit Function
    
    Dim wb As Excel.Workbook
    Set wb = GetObject(WorkbookPath)
    
    Dim Result As IExcelWorkbook
    Set Result = ExcelWorkbook.Create(wb)
    
    Set GetFrom = Result
    
End Function
Public Function OpenFrom( _
       ByVal WorkbookPath As String) As IExcelWorkbook
       
    If (Fso.FileExists(WorkbookPath) = False) Then _
        Exit Function
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    Dim wb As Excel.Workbook
    Set wb = Application.Workbooks.Open(WorkbookPath)
    
    Dim Result As IExcelWorkbook
    Set Result = ExcelWorkbook.Create(wb)
    
    Set OpenFrom = Result
        
End Function

Public Function Dispose()
    
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    
End Function
'---------------------------------------------
Private Function IExcelWorkbooksService_GetFrom(ByVal WorkbookPath As String) As IExcelWorkbook
    Set IExcelWorkbooksService_GetFrom = GetFrom(WorkbookPath)
End Function
Private Function IExcelWorkbooksService_OpenFrom(ByVal WorkbookPath As String) As IExcelWorkbook
    Set IExcelWorkbooksService_OpenFrom = OpenFrom(WorkbookPath)
End Function

Private Property Get IExcelWorkbooksService_ActiveBook() As IExcelWorkbook
    Set IExcelWorkbooksService_ActiveBook = ActiveBook
End Property
'---------------------------------------------
Private Sub IDisposableExcelWorbooksService_Dispose()
    Dispose
End Sub

