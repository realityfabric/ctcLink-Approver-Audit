Attribute VB_Name = "Tests_EA_Properties"
'@TestModule
'@Folder("Tests.Properties")

Option Explicit
Option Private Module

Private Assert As Object

'@ModuleInitialize
Private Sub ModuleInitialize()
    Set Assert = CreateObject("Rubberduck.AssertClass")
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
    Set Assert = Nothing
End Sub

'@TestMethod("Properties")
Private Sub FromChartfield_Numeric_NoLeadingZeroesOnLet()
    On Error GoTo TestFail
    
    'Arrange:
    Dim EA As ExpenseApproval
    Set EA = New ExpenseApproval
    
    'Act:
    EA.FromChartfield = "1"
    
    'Assert:
    Assert.AreEqual "00001", EA.FromChartfield

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Properties")
Private Sub FromChartfield_Numeric_LeadingZeroesOnLet()
    On Error GoTo TestFail
    
    'Arrange:
    Dim EA As ExpenseApproval
    Set EA = New ExpenseApproval
    
    'Act:
    EA.FromChartfield = "00001"
    
    'Assert:
    Assert.AreEqual "00001", EA.FromChartfield

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Properties")
Private Sub ToChartfield_Numeric_NoLeadingZeroesOnLet()
    On Error GoTo TestFail
    
    'Arrange:
    Dim EA As ExpenseApproval
    Set EA = New ExpenseApproval
    
    'Act:
    EA.ToChartfield = "1"
    
    'Assert:
    Assert.AreEqual "00001", EA.ToChartfield

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Properties")
Private Sub ToChartfield_Numeric_LeadingZeroesOnLet()
    On Error GoTo TestFail
    
    'Arrange:
    Dim EA As ExpenseApproval
    Set EA = New ExpenseApproval
    
    'Act:
    EA.ToChartfield = "00001"
    
    'Assert:
    Assert.AreEqual "00001", EA.ToChartfield

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub
