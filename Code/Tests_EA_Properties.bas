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

'@TestMethod("Properties")
Private Sub FirstName()
    On Error GoTo TestFail
    
    'Arrange:
    Dim EA As ExpenseApproval
    Set EA = New ExpenseApproval
    
    'Act:
    EA.FirstName = "John"
    
    'Assert:
    Assert.AreEqual "John", EA.FirstName

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Properties")
Private Sub LastName()
    On Error GoTo TestFail
    
    'Arrange:
    Dim EA As ExpenseApproval
    Set EA = New ExpenseApproval
    
    'Act:
    EA.LastName = "Doe"
    
    'Assert:
    Assert.AreEqual "Doe", EA.LastName

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Properties")
Private Sub DeptDesc()
On Error GoTo TestFail
    
    'Arrange:
    Dim EA As ExpenseApproval
    Set EA = New ExpenseApproval
    
    'Act:
    EA.DeptDesc = "Test Department Description"
    
    'Assert:
    Assert.AreEqual "Test Department Description", EA.DeptDesc

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Properties")
Private Sub BusinessUnit()
On Error GoTo TestFail
    
    'Arrange:
    Dim EA As ExpenseApproval
    Set EA = New ExpenseApproval
    
    'Act:
    EA.BusinessUnit = "WA000"
    
    'Assert:
    Assert.AreEqual "WA000", EA.BusinessUnit

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub
