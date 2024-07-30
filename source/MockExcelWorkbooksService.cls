VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MockExcelWorkbooksService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Tests.Mocks.Excel")
Option Explicit

Implements IExcelWorkbooksService
Implements IDisposableExcelWorbooksService

Public Property Get Self() As IExcelWorkbooksService
    Set Self = Me
End Property

Public Function GetFrom( _
       ByVal WorkbookPath As String) As IExcelWorkbook
       
Debug.Print
Debug.Print "Mock ExcelWorkbooksService.GetFrom" & _
            "   Path: "; WorkbookPath

    Set GetFrom = MockExcelWorkbook.Create()
    
End Function
Public Function OpenFrom( _
       ByVal WorkbookPath As String) As IExcelWorkbook
       
Debug.Print
Debug.Print "Mock ExcelWorkbooksService.OpenFrom" & _
            "   Path: "; WorkbookPath

    Set OpenFrom = MockExcelWorkbook.Create()
        
End Function
Public Function Dispose()

Debug.Print
Debug.Print "Mock ExcelWorkbooksService.Dispose"

End Function

Public Property Get ActiveBook() As IExcelWorkbook
    Set ActiveBook = MockExcelWorkbook.Create()
End Property
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

