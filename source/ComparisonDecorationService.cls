VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ComparisonDecorationService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "eco_transferToWord.Core.Comparison Results Decoration"
Option Explicit
Implements IComparisonDecorationService

Private Type TComparisonResultsDecorator
    
    Fso As IFileSystemObject
    Participants As List_IParticipant_
    
    TemporaryFolderPath As String
    Progress As IProgressService
    
    ParticipantTablesDecoratorFactory As IParcicipantTablesDecoratorFact
    ComparisonBlockDecoratorFactory As IComparisonBlockDecoratorFactor
    
    Workbooks As IExcelWorkbooksService

End Type

Private this As TComparisonResultsDecorator

Public Property Get Self() As IComparisonDecorationService
    Set Self = Me
End Property

Public Property Get Fso() As IFileSystemObject
    Set Fso = this.Fso
End Property
Public Property Set Fso(ByVal RHS As IFileSystemObject)
    Set this.Fso = RHS
End Property

Public Property Get Participants() As List_IParticipant_
    Set Participants = this.Participants
End Property
Public Property Set Participants(ByVal RHS As List_IParticipant_)
    Set this.Participants = RHS
End Property

Public Property Get TemporaryFolderPath() As String
    TemporaryFolderPath = this.TemporaryFolderPath
End Property
Public Property Let TemporaryFolderPath(ByVal RHS As String)
    this.TemporaryFolderPath = RHS
End Property

Public Property Get Progress() As IProgressService
    Set Progress = this.Progress
End Property
Public Property Set Progress(ByVal RHS As IProgressService)
    Set this.Progress = RHS
End Property

Public Property Get ParticipantTablesDecoratorFactory() As IParcicipantTablesDecoratorFact
    Set ParticipantTablesDecoratorFactory = this.ParticipantTablesDecoratorFactory
End Property
Public Property Set ParticipantTablesDecoratorFactory(ByVal RHS As IParcicipantTablesDecoratorFact)
    Set this.ParticipantTablesDecoratorFactory = RHS
End Property

Public Property Get ComparisonBlockDecoratorFactory() As IComparisonBlockDecoratorFactor
    Set ComparisonBlockDecoratorFactory = this.ComparisonBlockDecoratorFactory
End Property
Public Property Set ComparisonBlockDecoratorFactory(ByVal RHS As IComparisonBlockDecoratorFactor)
    Set this.ComparisonBlockDecoratorFactory = RHS
End Property

Public Property Get Workbooks() As IExcelWorkbooksService
    Set Workbooks = this.Workbooks
End Property
Public Property Set Workbooks(ByVal RHS As IExcelWorkbooksService)
    Set this.Workbooks = RHS
End Property

Public Sub DecorateComparisonResults()
    
    Dim Chartbook As IExcelWorkbook
    Set Chartbook = GetChartBook(ComparisonChartsStrategyFactory)
    
    Dim i As Long
    For i = 1 To Participants.Count
    
        Dim Model As IComparisonDataParticipant
        Set Model = Participants(i)
        
        Dim Files As List_IFile_
        Set Files = GetParticipantTemplates(Model.Number.Value)

        DecorateParticipantTable New MeasuredValuesTableStrategy, Files, i
        DecorateParticipantTable New ReferenceValuesTableStrategy, Files, i
        
        DecorateComparisonBlock New ComparisonTablesStrategyFactory, Files, i
        DecorateComparisonBlock ComparisonChartsStrategyFactory.Create(Chartbook), Files, i
        
        DecorateParticipantTable New ResultEvaluationTableStrategy, Files, i

    Next
    
    DisposeBooks Chartbook, Workbooks

End Sub
    Private Sub DisposeBooks( _
            ByVal Chartbook As ISavableExcelWorkbook, _
            ByVal Workbooks As IDisposableExcelWorbooksService)
        
        Chartbook.CloseWithSave
            
        Workbooks.Dispose
            
    End Sub
    Private Function GetChartBook( _
            ByVal Strategy As IComparisonTemplateFileStrategy) As IExcelWorkbook
        
        If (Fso Is Nothing = True) Then _
            Exit Function
        
        If (Workbooks Is Nothing = True) Then _
            Exit Function
        
        Dim Path As String
        Path = Fso.BuildPath(TemporaryFolderPath, Strategy.TemplateFileMask & ".xlsx")
        
        Dim Result As IExcelWorkbook
        Set Result = Workbooks.OpenFrom(Path)
        
        Set GetChartBook = Result
            
    End Function
    Private Function GetParticipantTemplates( _
            ByVal ParticipantNumber As String) As List_IFile_

        Dim Destination As String
        Destination = Fso.BuildPath(TemporaryFolderPath, ParticipantNumber)

        Dim Result As List_IFile_
        Set Result = Fso.GetFolderFiles(Destination)

        Set GetParticipantTemplates = Result

    End Function
    
    Private Sub DecorateParticipantTable( _
            ByVal Strategy As IComparisonDecorationStrategy, _
            ByVal Files As List_IFile_, _
            ByVal ExecutionCount As Long)

        ShowProgress Strategy, ExecutionCount

        Dim Model As IResultsObtainedParticipant
        Set Model = Participants(ExecutionCount)

        Dim Decorator As IParticipantTablesDecorator
        Set Decorator = ParticipantTablesDecoratorFactory.Create(Strategy, Files, Model.Results)
    
        Decorator.ExecuteFilling
        Decorator.SaveChanges

    End Sub
        Private Sub ShowProgress( _
            ByVal Strategy As IComparisonProgressNotifyStrate, _
            ByVal ExecutionCount As Long)
            
            Dim Header As String
            Header = GetProgressHeader(ExecutionCount)
            
            Progress.Show Header, Strategy.ProgressFooter
            
        End Sub
            Private Function GetProgressHeader( _
                    ByVal ExecutionCount As Long) As String
                
                Dim Result As String
                Result = "Результаты участника " & " (" & ExecutionCount & " / " & Participants.Count & ")"
                
                GetProgressHeader = Result
                
            End Function
            
    Private Sub DecorateComparisonBlock( _
            ByVal Strategy As IComparisonBlockStrategyFactory, _
            ByVal Files As List_IFile_, _
            ByVal ExecutionCount As Long)
    
        Dim Decorator As IComparisonBlockDecorator
        Set Decorator = ComparisonBlockDecoratorFactory.Create(Strategy, Files)
        
        Dim BlockCount As Long
        
        Dim Model As IResultsObtainedParticipant
        Set Model = Participants(ExecutionCount)
        
        Dim i As Long
        For i = Model.Results.Count To 1 Step -1
        
            BlockCount = BlockCount + 1
            '---------------------------------------------
            Dim Header As String
            Header = GetProgressHeader(ExecutionCount)
            
            Dim Footer As String
            Footer = GetProgressFooter(Strategy, BlockCount, Model.Results.Count)
            
            Progress.Show Header, Footer
            '---------------------------------------------
            Dim ParticipantNumber As String
            ParticipantNumber = GetNumber(Participants(ExecutionCount))
            
            Decorator.ExecuteFilling ParticipantNumber, i
            
        Next
        
        Decorator.Dispose
        Decorator.SaveChanges
                    
    End Sub
    Private Function GetNumber( _
            ByVal Model As IComparisonDataParticipant) As String
             
        GetNumber = Model.Number.Value
          
    End Function
    
    
    Private Function GetProgressFooter( _
                ByVal Strategy As IComparisonProgressNotifyStrate, _
                ByVal BlockExecutionCount As Long, _
                ByVal BlockTotalCount As Long) As String
            
            Dim Footer As String
            Footer = Strategy.ProgressFooter & ". Блок " & BlockExecutionCount & " / " & BlockTotalCount
            
            GetProgressFooter = Footer
            
        End Function
'---------------------------------------------
Private Sub IComparisonDecorationService_DecorateComparisonResults()
    DecorateComparisonResults
End Sub

