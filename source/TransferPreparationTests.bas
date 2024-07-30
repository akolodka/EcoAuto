Attribute VB_Name = "TransferPreparationTests"
'@TestModule
'@Folder("Tests")

Option Explicit
Option Private Module

#Const LateBind = LateBindTests

#If LateBind Then
    Private Assert As Object
    Private Fakes As Object
#Else
    Private Assert As Rubberduck.AssertClass
    Private Fakes As Rubberduck.FakesProvider
#End If

'@ModuleInitialize
Private Sub ModuleInitialize()
    'this method runs once per module.
#If LateBind Then
        Set Assert = CreateObject("Rubberduck.AssertClass")
        Set Fakes = CreateObject("Rubberduck.FakesProvider")
#Else
        Set Assert = New Rubberduck.AssertClass
        Set Fakes = New Rubberduck.FakesProvider
#End If
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
    'this method runs once per module.
    Set Assert = Nothing
    Set Fakes = Nothing
End Sub

'@TestInitialize
Private Sub TestInitialize()
    'This method runs before every test in the module..
End Sub

'@TestCleanup
Private Sub TestCleanup()
    'this method runs after every test in the module.
End Sub

'@TestMethod("Uncategorized")
Private Sub MockPreparationTest()
    On Error GoTo TestFail
    
    'Arrange:
    Dim Initial As IInitializationService
    Set Initial = MockInitializationService.Create()
    
    Dim Model As ITransferMenuModel
    Set Model = MockTransferMenuModel.Create(Initial)
    
    Dim PreparerFactory As ITransferPreparationServiceFact
    Set PreparerFactory = TransferPreparationServiceFacto.Create(Initial)
    
    Dim Preparer As ITransferPreparationService
    Set Preparer = PreparerFactory.Create(Model)
        
    Dim DecorationFactory As IComparisonDecorationServiceFac
    Set DecorationFactory = ComparisonDecorationServiceFact.Create(Initial)
    
    Dim Decorator As IComparisonDecorationService
    Set Decorator = DecorationFactory.Create(Model)
    
    'Act:
    Preparer.PrepareTemplates
    Decorator.DecorateComparisonResults
    
    Preparer.Dispose
    
    'Assert:
    Assert.Succeed
    

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Uncategorized")
Private Sub BattlePreparationTest()
    On Error GoTo TestFail
    
    'Arrange:
    Dim Initial As ITransferProcessInitialization
    Set Initial = MockInitializationService.CreateBattleSilent()
    
    Dim Model As ITransferMenuModel
    Set Model = MockTransferMenuModel.Create(Initial)
    
    Dim PreparerFactory As ITransferPreparationServiceFact
    Set PreparerFactory = TransferPreparationServiceFacto.Create(Initial)
    
    Dim Preparer As ITransferPreparationService
    Set Preparer = PreparerFactory.Create(Model)
    
    Dim DecorationFactory As IComparisonDecorationServiceFac
    Set DecorationFactory = ComparisonDecorationServiceFact.Create(Initial)
    
    Dim Decorator As IComparisonDecorationService
    Set Decorator = DecorationFactory.Create(Model)
    
    'Act:
    Preparer.PrepareTemplates
    Decorator.DecorateComparisonResults
    
   ' Preparer.Dispose

    'Assert:
    Assert.Succeed
    
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    Initial.Word.Dispose
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

