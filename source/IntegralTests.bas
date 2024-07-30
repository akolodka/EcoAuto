Attribute VB_Name = "IntegralTests"
'@TestModule
'@Folder "Tests"
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
    'This method runs before every test in the module.
End Sub

'@TestCleanup
Private Sub TestCleanup()
    'this method runs after every test in the module.
End Sub

'@TestMethod("Initialization")
Private Sub IntegralMockTest()
    On Error GoTo TestFail

    'Arrange:
    Dim Initial As IInitializationService
    Set Initial = MockInitializationService.Create()
    
    Dim Dialog As ITransferDialogAction
    Set Dialog = MockTransferMenuPresenter.Create(Initial)
    
    'Act:
    Dialog.InitiateTransfer

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

'@TestMethod("Initialization")
Private Sub IntegralBattleSilentTest()
    On Error GoTo TestFail

    'Arrange:
    Dim Initial As IInitializationService
    Set Initial = MockInitializationService.CreateBattleSilent()
    
    Dim Dialog As ITransferDialogAction
    Set Dialog = MockTransferMenuPresenter.Create(Initial)
    
    'Act:
    Dialog.InitiateTransfer

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

