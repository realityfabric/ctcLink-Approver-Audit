Attribute VB_Name = "Tests_Employee"
'@TestModule
'@Folder("Tests")


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

'@TestMethod("Wrapper")
Private Sub AddRole_NewRole()
    On Error GoTo TestFail
    
    'Arrange:
    Dim Emp As Employee
    Set Emp = New Employee
    
    'Act:
    Emp.AddRole ("ZZ Test Role")
    
    'Assert:
    Assert.IsTrue "ZZ Test Role" = Emp.Role(1)

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Wrapper")
Private Sub AddRole_ExistingRole()
    On Error GoTo TestFail
    
    'Arrange:
    Dim Emp As Employee
    Set Emp = New Employee
    
    'Act:
    Emp.AddRole ("ZZ Test Role")
    Emp.AddRole ("ZZ Test Role")
    
    'Assert:
    Assert.IsTrue "ZZ Test Role" = Emp.Role(1)
    Assert.IsTrue Emp.RoleCount = 1

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Uncategorized")
Private Sub HasRole_False()
    On Error GoTo TestFail
    
    'Arrange:
    Dim Emp As Employee
    Set Emp = New Employee
    
    'Act:
    'Assert:
    Assert.IsFalse Emp.HasRole("HasRole Test Role")
    
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Uncategorized")
Private Sub HasRole_True()
    On Error GoTo TestFail
    
    'Arrange:
    Dim Emp As Employee
    Set Emp = New Employee
    Emp.AddRole "HasRole Test Role"
    
    'Act:
    'Assert:
    Assert.IsTrue Emp.HasRole("HasRole Test Role")
    
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub
