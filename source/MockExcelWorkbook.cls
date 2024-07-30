VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MockExcelWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder("Tests.Mocks.Excel.Workbook")
'@PredeclaredId
Option Explicit

Implements IExcelWorkbook
Implements ISavableExcelWorkbook

Public Function Create() As IExcelWorkbook
    
    With New MockExcelWorkbook
        Set Create = .Self
    End With

End Function

Public Property Get Self() As IExcelWorkbook
    Set Self = Me
End Property

Private Property Get Worksheets( _
        ByVal Name As String) As IExcelWorksheet
        
Debug.Print
Debug.Print "Mock ExcelWorkbook.Worksheets" & _
            "   Name: "; Name

    Set Worksheets = MockExcelWorksheet.Create()

End Property

Private Property Get FirstSheet() As IExcelWorksheet
        
Debug.Print
Debug.Print "Mock ExcelWorkbook.FirstSheet"

    Set FirstSheet = MockExcelWorksheet.Create()

End Property

Private Property Get LastSheet() As IExcelWorksheet
        
Debug.Print
Debug.Print "Mock ExcelWorkbook.LastSheet"

    Set LastSheet = MockExcelWorksheet.Create()

End Property
Private Sub CloseWithSave()

Debug.Print
Debug.Print "Mock ExcelWorkbook.CloseWithSave"

End Sub

Private Function IsSheetExists( _
        ByVal Name As String) As Boolean
        
    Const Result As Boolean = True
    
Debug.Print
Debug.Print "Mock ExcelWorkbook.IsSheetExists" & _
            "   Name: "; Name; _
            "   Result: "; Result
            
        
        IsSheetExists = Result
        
End Function
'---------------------------------------------
Private Property Get IExcelWorkbook_Worksheets(ByVal Name As String) As IExcelWorksheet
    Set IExcelWorkbook_Worksheets = Worksheets(Name)
End Property
Private Property Get IExcelWorkbook_LastSheet() As IExcelWorksheet
    Set IExcelWorkbook_LastSheet = LastSheet
End Property
Private Property Get IExcelWorkbook_FirstSheet() As IExcelWorksheet
    Set IExcelWorkbook_FirstSheet = FirstSheet
End Property
Private Function IExcelWorkbook_IsSheetExist(ByVal Name As String) As Boolean
    IExcelWorkbook_IsSheetExist = IsSheetExists(Name)
End Function
'---------------------------------------------
Private Sub ISavableExcelWorkbook_CloseWithSave()
    CloseWithSave
End Sub


