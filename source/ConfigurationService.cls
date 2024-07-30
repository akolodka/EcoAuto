VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ConfigurationService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder("eco_transferToWord.Core.Configuration")
'@PredeclaredId
Option Explicit

Public Function Create() As IConfiguration
    
    Dim Fso As FileSystemObject
    Set Fso = New FileSystemObject
    
    With New ConfigurationDTO
        
        .TemporaryFolderPath = Fso.BuildPath(ActiveWorkbook.Path, "temporary")
        .ResultFolderPath = Fso.BuildPath(ActiveWorkbook.Path, "result")
        
        .TemplatesFolderPath = Fso.BuildPath(ThisWorkbook.Path, "templates")
        .StaticDataFolderPath = Fso.BuildPath(ThisWorkbook.Path, "static")
        .SubcontractFolderPath = Fso.BuildPath(ThisWorkbook.Path, "subcontract")
        .ComparisonAuxillaryFolderPath = Fso.BuildPath(ThisWorkbook.Path, "comparison")
        
        .WorksheetName = "Результаты участников"
        
        .MeasuredValuesRangesWorkbookPath = Fso.BuildPath(ThisWorkbook.Path, "measuredValuesRanges.xlsx")
        .ComparisonChartsTemplateWorkbookPath = Fso.BuildPath(ThisWorkbook.Path, "comparisonCharts.xlsx")
        
        Set Create = .Self
        
    End With
    
End Function

