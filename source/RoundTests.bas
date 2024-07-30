Attribute VB_Name = "RoundTests"
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

'@TestMethod("ComparisonDecoration")
Private Sub RounderGetValueTest()
    On Error GoTo TestFail
    
    'Arrange:
    Dim Rounder As IValueRounder
    Set Rounder = ValueRounder.Create(100.01231, 1.223232)
    'Act:
    Debug.Print Rounder.GetValue
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

'@TestMethod("ComparisonDecoration")
Private Sub RounderGetUncertaintyTest()
    On Error GoTo TestFail
    
    'Arrange:
    Dim Rounder As IValueRounder
    Set Rounder = ValueRounder.Create(100.01231, 1.223232)
    'Act:
    Debug.Print Rounder.GetUncertainty
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

'@TestMethod("ComparisonDecoration")
Private Sub RounderGetRoundedToHundredthsTest()
    On Error GoTo TestFail
    
    'Arrange:
    Dim Rounder As IValueRounder
    Set Rounder = ValueRounder.Create(100.01231, 1.223232)
    'Act:
    Debug.Print
    Debug.Print Rounder.GetRoundedToHundredths
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
