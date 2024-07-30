VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ConfigurationDTO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "eco_transferToWord.Core.Configuration"
Option Explicit

Implements IConfiguration
Implements ITransferPreparationConfig
Implements IResultPresentationConfiguratio

Private Type TConfigurationService
    
    TemporaryFolderPath As String
    ResultFolderPath As String
    
    TemplatesFolderPath As String
    StaticDataFolderPath As String
    SubcontractFolderPath As String
    ComparisonAuxillaryFolderPath As String
    
    WorksheetName As String
    
    MeasuredValuesRangesWorkbookPath As String
    ComparisonChartsTemplateWorkbookPath As String
    
End Type

Private this As TConfigurationService

Public Property Get Self() As IConfiguration
    Set Self = Me
End Property

Public Property Get TemporaryFolderPath() As String
    TemporaryFolderPath = this.TemporaryFolderPath
End Property
Public Property Let TemporaryFolderPath(ByVal RHS As String)
    this.TemporaryFolderPath = RHS
End Property

Public Property Get ResultFolderPath() As String
    ResultFolderPath = this.ResultFolderPath
End Property
Public Property Let ResultFolderPath(ByVal RHS As String)
    this.ResultFolderPath = RHS
End Property

Public Property Get TemplatesFolderPath() As String
    TemplatesFolderPath = this.TemplatesFolderPath
End Property
Public Property Let TemplatesFolderPath(ByVal RHS As String)
    this.TemplatesFolderPath = RHS
End Property

Public Property Get StaticDataFolderPath() As String
    StaticDataFolderPath = this.StaticDataFolderPath
End Property
Public Property Let StaticDataFolderPath(ByVal RHS As String)
    this.StaticDataFolderPath = RHS
End Property

Public Property Get SubcontractFolderPath() As String
    SubcontractFolderPath = this.SubcontractFolderPath
End Property
Public Property Let SubcontractFolderPath(ByVal RHS As String)
    this.SubcontractFolderPath = RHS
End Property

Public Property Get ComparisonAuxillaryFolderPath() As String
    ComparisonAuxillaryFolderPath = this.ComparisonAuxillaryFolderPath
End Property
Public Property Let ComparisonAuxillaryFolderPath(ByVal RHS As String)
    this.ComparisonAuxillaryFolderPath = RHS
End Property

Public Property Get WorksheetName() As String
    WorksheetName = this.WorksheetName
End Property
Public Property Let WorksheetName(ByVal RHS As String)
    this.WorksheetName = RHS
End Property

Public Property Get MeasuredValuesRangesWorkbookPath() As String
    MeasuredValuesRangesWorkbookPath = this.MeasuredValuesRangesWorkbookPath
End Property
Public Property Let MeasuredValuesRangesWorkbookPath(ByVal RHS As String)
    this.MeasuredValuesRangesWorkbookPath = RHS
End Property

Public Property Get ComparisonChartsTemplateWorkbookPath() As String
    ComparisonChartsTemplateWorkbookPath = this.ComparisonChartsTemplateWorkbookPath
End Property
Public Property Let ComparisonChartsTemplateWorkbookPath(ByVal RHS As String)
    this.ComparisonChartsTemplateWorkbookPath = RHS
End Property
'---------------------------------------------
Private Property Get IConfiguration_TemporaryFolderPath() As String
    IConfiguration_TemporaryFolderPath = TemporaryFolderPath
End Property
Private Property Get IConfiguration_ResultFolderPath() As String
    IConfiguration_ResultFolderPath = ResultFolderPath
End Property
Private Property Get IConfiguration_TemplatesFolderPath() As String
    IConfiguration_TemplatesFolderPath = TemplatesFolderPath
End Property
Private Property Get IConfiguration_StaticDataFolderPath() As String
    IConfiguration_StaticDataFolderPath = StaticDataFolderPath
End Property
Private Property Get IConfiguration_ComparisonAuxillaryFolderPath() As String
    IConfiguration_ComparisonAuxillaryFolderPath = ComparisonAuxillaryFolderPath
End Property
Private Property Get IConfiguration_SubcontractFolderPath() As String
    IConfiguration_SubcontractFolderPath = SubcontractFolderPath
End Property
Private Property Get IConfiguration_WorksheetName() As String
    IConfiguration_WorksheetName = WorksheetName
End Property
Private Property Get IConfiguration_MeasuredValuesRangesWorkbookPath() As String
    IConfiguration_MeasuredValuesRangesWorkbookPath = MeasuredValuesRangesWorkbookPath
End Property
'---------------------------------------------
Private Property Get ITransferPreparationConfig_DocumentTemplatesFolderPath() As String
    ITransferPreparationConfig_DocumentTemplatesFolderPath = TemplatesFolderPath
End Property
Private Property Get ITransferPreparationConfig_AuxillaryTemplatesFolderPath() As String
    ITransferPreparationConfig_AuxillaryTemplatesFolderPath = ComparisonAuxillaryFolderPath
End Property
Private Property Get ITransferPreparationConfig_TemporaryFolderPath() As String
    ITransferPreparationConfig_TemporaryFolderPath = TemporaryFolderPath
End Property
Private Property Get ITransferPreparationConfig_ResultFolderPath() As String
    ITransferPreparationConfig_ResultFolderPath = ResultFolderPath
End Property
Private Property Get ITransferPreparationConfig_MeasuredValuesRangesWorkbookPath() As String
    ITransferPreparationConfig_MeasuredValuesRangesWorkbookPath = MeasuredValuesRangesWorkbookPath
End Property
Private Property Get ITransferPreparationConfig_ComparisonChartsTemplateWorkbookPath() As String
    ITransferPreparationConfig_ComparisonChartsTemplateWorkbookPath = ComparisonChartsTemplateWorkbookPath
End Property
'---------------------------------------------
Private Property Get IResultPresentationConfiguratio_TemporaryFolderPath() As String
    IResultPresentationConfiguratio_TemporaryFolderPath = TemporaryFolderPath
End Property
Private Property Get IResultPresentationConfiguratio_ResultFolderPath() As String
    IResultPresentationConfiguratio_ResultFolderPath = ResultFolderPath
End Property

