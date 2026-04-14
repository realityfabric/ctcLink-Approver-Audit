Attribute VB_Name = "TestModule_Employee_Integration"
' Audit security roles and configurations for Approvals in ctcLink.
'    Copyright (C) 2026 Jessica Fairchild aka Jessica Jones-Copeland
'
'    This program is free software: you can redistribute it and/or modify
'    it under the terms of the GNU General Public License as published by
'    the Free Software Foundation, either version 3 of the License, or
'    (at your option) any later version.
'
'    This program is distributed in the hope that it will be useful,
'    but WITHOUT ANY WARRANTY; without even the implied warranty of
'    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'    GNU General Public License for more details.
'
'    You should have received a copy of the GNU General Public License
'    along with this program.  If not, see <https://www.gnu.org/licenses/>.

'@TestModule
'@Folder("Tests.Integrations")

Option Explicit
Option Private Module

Private Assert As Object
Private ActiveEmployee As Employee
Private InactiveEmployee As Employee

'@ModuleInitialize
Private Sub ModuleInitialize()
    Set Assert = CreateObject("Rubberduck.AssertClass")
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
    Set Assert = Nothing
End Sub

'@TestInitialize
Private Sub TestInitialize()
    Set ActiveEmployee = New Employee
    Set InactiveEmployee = New Employee

    ActiveEmployee.EmplID = "1"
    ActiveEmployee.Name = "Doe, Jane"
    ActiveEmployee.HRStatus = "A"

    InactiveEmployee.EmplID = "2"
    InactiveEmployee.Name = "Smith, John"
    InactiveEmployee.HRStatus = "I"
End Sub

'@TestCleanup
Private Sub TestCleanup()
    Set ActiveEmployee = Nothing
    Set InactiveEmployee = Nothing
End Sub

'@TestMethod("No Fail")
Private Sub AddDepartment_ActiveEmplIDMatch_NoFail()
    On Error GoTo TestFail
    
    'Arrange:
    Dim Dept As Department
    Set Dept = New Department
    Dept.DeptID = "TEST"
    Dept.ManagerID = "1"
    Dept.Description = "Test Department"
    
    'Act:
    ActiveEmployee.AddDepartment Dept
    
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

'@TestMethod("No Fail")
Private Sub AddDepartment_InactiveEmplIDMatch_NoFail()
    On Error GoTo TestFail
    
    'Arrange:
    Dim Dept As Department
    Set Dept = New Department
    Dept.DeptID = "TEST"
    Dept.ManagerID = "2"
    Dept.Description = "Test Department"
    
    'Act:
    InactiveEmployee.AddDepartment Dept
    
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

'@TestMethod("No Fail")
Private Sub AddDepartment_ActiveEmplIDNoMatch_NoFail()
    On Error GoTo TestFail
    
    'Arrange:
    Dim Dept As Department
    Set Dept = New Department
    Dept.DeptID = "TEST"
    Dept.ManagerID = "2"
    Dept.Description = "Test Department"
    
    'Act:
    ActiveEmployee.AddDepartment Dept
    
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
Private Sub AddDepartment_ActiveEmplIDMatch_CountCorrect()
    On Error GoTo TestFail
    
    'Arrange:
    Dim Dept As Department
    Set Dept = New Department
    Dept.DeptID = "TEST"
    Dept.ManagerID = "1"
    Dept.Description = "Test Department"
    
    'Act:
    ActiveEmployee.AddDepartment Dept
    
    'Assert:
    Assert.IsTrue ActiveEmployee.DepartmentCount = 1

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Uncategorized")
Private Sub AddDepartment_ActiveEmplIDNoMatch_CountCorrect()
    On Error GoTo TestFail
    
    'Arrange:
    Dim Dept As Department
    Set Dept = New Department
    Dept.DeptID = "TEST"
    Dept.ManagerID = "2"
    Dept.Description = "Test Department"
    
    'Act:
    ActiveEmployee.AddDepartment Dept
    
    'Assert:
    Assert.IsTrue ActiveEmployee.DepartmentCount = 0

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("No Fail")
Private Sub GetDepartment_DeptExists_NoFail()
    On Error GoTo TestFail
    
    'Arrange:
    Dim Dept As Department
    Set Dept = New Department
    Dept.DeptID = "TEST"
    Dept.ManagerID = "1"
    Dept.Description = "Test Department"
    ActiveEmployee.AddDepartment Dept
    
    'Act:
    '@Ignore FunctionReturnValueDiscarded
    ActiveEmployee.Department 1
    
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
Private Sub AddExpenseApproval_ActiveEmplIDMatch_CountCorrect()
    On Error GoTo TestFail
    
    'Arrange:
    Dim ExpAppr As ExpenseApproval
    Set ExpAppr = New ExpenseApproval
    ExpAppr.ApproverType = "EXAPPROVER"
    ExpAppr.BusinessUnit = "WA999"
    ExpAppr.EmplID = "1"
    ExpAppr.FirstName = "Jane"
    ExpAppr.LastName = "Doe"
    ExpAppr.FromChartfield = "10"
    ExpAppr.ToChartfield = "20"
    
    'Act:
    ActiveEmployee.AddExpenseApproval ExpAppr
    
    'Assert:
    Assert.IsTrue ActiveEmployee.ExpenseApprovalCount = 1

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

