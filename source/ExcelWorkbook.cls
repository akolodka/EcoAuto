VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ExcelWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder("eco_transferToWord.Core.Excel.Workbook")
'@PredeclaredId
Option Explicit

Implements IExcelWorkbook
Implements ISavableExcelWorkbook

Private Type TExcelWorkbook
    Workbook As Excel.Workbook
End Type

Private this As TExcelWorkbook

Public Function Create( _
       ByVal Workbook As Excel.Workbook) As IExcelWorkbook
    
    With New ExcelWorkbook
    
        Set .Workbook = Workbook
        Set Create = .Self
        
    End With

End Function

Public Property Get Workbook() As Excel.Workbook
    Set Workbook = this.Workbook
End Property
Public Property Set Workbook(ByVal RHS As Excel.Workbook)
    Set this.Workbook = RHS
End Property

Public Property Get Self() As IExcelWorkbook
    Set Self = Me
End Property

Public Property Get Worksheets( _
        ByVal Name As String) As IExcelWorksheet
    
    Dim isExist As Boolean
    isExist = IsSheetExist(Name)
    
    If (isExist = False) Then _
        Exit Property
    
    Dim Ws As Excel.Worksheet
    Set Ws = Workbook.Worksheets(Name)
    
    Dim Result As IExcelWorksheet
    Set Result = ExcelWorksheet.Create(Ws)
    
    Set Worksheets = Result

End Property

Public Property Get FirstSheet() As IExcelWorksheet
    
    Dim Ws As Excel.Worksheet
    Set Ws = Workbook.Worksheets(1)
    
    Dim Result As IExcelWorksheet
    Set Result = ExcelWorksheet.Create(Ws)
    
    Set FirstSheet = Result

End Property

Public Property Get LastSheet() As IExcelWorksheet
    
    Dim Ws As Excel.Worksheet
    Set Ws = Workbook.Worksheets(Workbook.Worksheets.Count)
    
    Dim Result As IExcelWorksheet
    Set Result = ExcelWorksheet.Create(Ws)
    
    Set LastSheet = Result

End Property
Private Function IsSheetExist( _
        ByVal Name As String) As Boolean
        
    Dim i As Long
    For i = 1 To Workbook.Worksheets.Count
        
        Dim Ws As Excel.Worksheet
        Set Ws = Workbook.Worksheets(i)
        
        If Ws.Name = Name Then
        
            IsSheetExist = True
            Exit Function
            
        End If
        
    Next
        
End Function
'---------------------------------------------
Private Property Get IExcelWorkbook_Worksheets(ByVal Name As String) As IExcelWorksheet
    Set IExcelWorkbook_Worksheets = Worksheets(Name)
End Property
Private Function IExcelWorkbook_IsSheetExist(ByVal Name As String) As Boolean
    IExcelWorkbook_IsSheetExist = IsSheetExist(Name)
End Function
Private Property Get IExcelWorkbook_LastSheet() As IExcelWorksheet
    Set IExcelWorkbook_LastSheet = LastSheet
End Property
Private Property Get IExcelWorkbook_FirstSheet() As IExcelWorksheet
    Set IExcelWorkbook_FirstSheet = FirstSheet
End Property
'---------------------------------------------
Private Sub ISavableExcelWorkbook_CloseWithSave()
    Workbook.Close True
End Sub



