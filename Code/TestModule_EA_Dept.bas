Attribute VB_Name = "TestModule_EA_Dept"
'@TestModule
'@Folder("Tests.Integrations")


Option Explicit
Option Private Module

Private Assert As Object
Private Fakes As Object

'@ModuleInitialize
Private Sub ModuleInitialize()
    'this method runs once per module.
    Set Assert = CreateObject("Rubberduck.AssertClass")
    Set Fakes = CreateObject("Rubberduck.FakesProvider")
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

'@TestMethod("Compare")
Private Sub TestMethod_EAMatchesDept_NotFromOrToField()
    On Error GoTo TestFail
    
    'Arrange:
    '@Ignore UseMeaningfulName
    Dim EA As ExpenseApproval
    Dim Dept As Department
    
    Set EA = New ExpenseApproval
    Set Dept = New Department
    
    EA.ApproverType = "EXAPPROVAL"
    EA.EmplID = "1"
    EA.FromChartfield = "1"
    EA.ToChartfield = "5"
    Dept.ManagerID = "1"
    Dept.DeptID = "3"
    
    'Act:
    'Assert:
    Assert.IsFalse EA.DepartmentManagerMismatch(Dept)
    Assert.IsTrue EA.DepartmentInRange(Dept.DeptID)

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Compare")
Private Sub TestMethod_EADeptMismatch_NotFromOrToField()
    On Error GoTo TestFail
    
    'Arrange:
    '@Ignore UseMeaningfulName
    Dim EA As ExpenseApproval
    Dim Dept As Department
    
    Set EA = New ExpenseApproval
    Set Dept = New Department
    
    EA.ApproverType = "EXAPPROVER"
    EA.EmplID = "1"
    EA.FromChartfield = "1"
    EA.ToChartfield = "5"
    Dept.ManagerID = "2"
    Dept.DeptID = "3"
    
    'Act:
    'Assert:
    Assert.IsTrue EA.DepartmentManagerMismatch(Dept)

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Compare")
Private Sub TestMethod_EADeptMismatch_FromField()
    On Error GoTo TestFail
    
    'Arrange:
    '@Ignore UseMeaningfulName
    Dim EA As ExpenseApproval
    Dim Dept As Department
    
    Set EA = New ExpenseApproval
    Set Dept = New Department
    
    EA.ApproverType = "EXAPPROVER"
    EA.EmplID = "1"
    EA.FromChartfield = "1"
    EA.ToChartfield = "5"
    Dept.ManagerID = "2"
    Dept.DeptID = "1"
    
    'Act:
    'Assert:
    Assert.IsTrue EA.DepartmentManagerMismatch(Dept)

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub
