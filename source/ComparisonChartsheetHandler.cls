VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ComparisonChartsheetHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "eco_transferToWord.Core.Comparison Results Decoration.Comparison Block Decoration.Chart Sheet Handler"
Option Explicit
Implements IComparisonChartsheetHandler

Private Type TComparisonChartsheetHandler
    
    Chartbook As IExcelWorkbook
    Model As List_IOutputResultRowModel_
    
End Type

Private this As TComparisonChartsheetHandler

Public Property Get Self() As IComparisonChartsheetHandler
    Set Self = Me
End Property

Public Property Get Chartbook() As IExcelWorkbook
    Set Chartbook = this.Chartbook
End Property
Public Property Set Chartbook(ByVal RHS As IExcelWorkbook)
    Set this.Chartbook = RHS
End Property

Public Property Get Model() As List_IOutputResultRowModel_
    Set Model = this.Model
End Property
Public Property Set Model(ByVal RHS As List_IOutputResultRowModel_)
    Set this.Model = RHS
End Property

Public Property Get IsSheetFilled() As Boolean
    
    If (Model.Count = vbEmpty) Then _
        Exit Property
    
    Dim BlockModel As IOutputResultsRowModel
    Set BlockModel = Model(1)
    
    Dim Result As Boolean
    Result = (Chartbook.IsSheetExist(BlockModel.BlockIndex) = True)
    
    IsSheetFilled = Result
    
End Property

Public Sub FillSheet()
    
    CopyTemplateSheet
    FillTemplateSheet
    
End Sub
    Private Sub CopyTemplateSheet()
        
        Dim Sheet As IClonableExcelWorksheet
        Set Sheet = Chartbook.LastSheet
        
        Sheet.CopyBefore Chartbook.FirstSheet
        
    End Sub
    Private Sub FillTemplateSheet()
        
        Dim Sheet As IExcelWorksheet
        Set Sheet = Chartbook.FirstSheet
        
        FillSheetName Sheet, Model(1)
        FillHeaderData Sheet, Model(1)
        
        Dim i As Long
        For i = Model.Count To 1 Step -1
            
            FillRow Sheet, Model(i)
            
        Next
        
        DeleteTemplateRow Sheet
        RefreshChartScale Sheet
        
    End Sub
        Private Sub FillSheetName( _
                ByVal Sheet As IEditableExcelWorksheet, _
                ByVal Model As IOutputResultsRowModel)

            Sheet.Name = Model.BlockIndex
            
        End Sub
        Private Sub FillHeaderData( _
                ByVal Sheet As IEditableExcelWorksheet, _
                ByVal Model As IComparisonChartSheetHeaderMode)

            Sheet.Cells(2, 5).Value = Model.ValueName
               
            Sheet.Cells(4, 6).Value = CDbl(Model.ReferenceValue)
            Sheet.Cells(5, 6).Value = CDbl(Model.ReferenceUncertainty)
            
        End Sub
        Private Sub FillRow( _
            ByVal Sheet As IEditableExcelWorksheet, _
            ByVal Model As IComparisonChartSheetRowModel)
            
            If (Model.Value = vbNullString) Then _
                Exit Sub
            
            Sheet.CloneRowDown 9
            
            Sheet.Cells(10, 2).Value = Model.ParticipantName
            Sheet.Cells(10, 3).Value = CDbl(Model.Value)
            Sheet.Cells(10, 4).Value = CDbl(Model.Uncertainty)
            
        End Sub
        Private Sub DeleteTemplateRow( _
                ByVal Sheet As IEditableExcelWorksheet)
            
            Sheet.DeleteRow 9
            
        End Sub
        Private Sub RefreshChartScale( _
                ByVal Sheet As IEditableExcelWorksheet)

            Dim TargetCell As IExcelRange
            Set TargetCell = Sheet.Cells(4, 8)
            
            If (TargetCell Is Nothing = True) Then _
                Exit Sub
                
            TargetCell.Calculate
            
            If (IsNumeric(TargetCell.Value) = False) Then _
                Exit Sub
            
            Dim MinValue As Double
            MinValue = TargetCell.Value
            
            Set TargetCell = Sheet.Cells(5, 8)
            
            If (TargetCell Is Nothing = True) Then _
                Exit Sub
            
            TargetCell.Calculate
            
            If (IsNumeric(TargetCell.Value) = False) Then _
                Exit Sub

            Dim MaxValue As Double
            MaxValue = TargetCell.Value
            
            Dim Chart As IExcelChart
            Set Chart = GetChart(Sheet)

            Chart.ApplyScale MinValue, MaxValue
        
        End Sub
                Private Function GetChart( _
                        ByVal Sheet As IChartableExcelWorksheet) As IExcelChart
                                
                    Set GetChart = Sheet.GetChart()
                                
                End Function
'---------------------------------------------
Private Property Get IComparisonChartsheetHandler_IsSheetFilled() As Boolean
    IComparisonChartsheetHandler_IsSheetFilled = IsSheetFilled
End Property

Private Sub IComparisonChartsheetHandler_FillSheet()
    FillSheet
End Sub

