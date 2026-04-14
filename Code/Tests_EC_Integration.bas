Attribute VB_Name = "Tests_EC_Integration"
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

'@ModuleInitialize
Private Sub ModuleInitialize()
    Set Assert = CreateObject("Rubberduck.AssertClass")
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
    Set Assert = Nothing
End Sub

'@TestMethod("Uncategorized")
Private Sub SetEmployeeDepartments_NoEmployeesNoDepartments()
    On Error GoTo TestFail
    
    'Arrange:
    Dim Employees As EmployeeCollection
    Dim Departments As DepartmentCollection
    
    Set Employees = New EmployeeCollection
    Set Departments = New DepartmentCollection
    
    'Act:
    Employees.SetEmployeeDepartments Departments
    
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
Private Sub SetEmployeeDepartments_OneEmployeesNoDepartments_CorrectCount()
    On Error GoTo TestFail
    
    'Arrange:
    Dim Employees As EmployeeCollection
    Dim Departments As DepartmentCollection
    Dim Emp As Employee
    
    Set Employees = New EmployeeCollection
    Set Departments = New DepartmentCollection
    Set Emp = New Employee
    
    Emp.EmplID = "1"
    Employees.Add Emp
    
    'Act:
    Employees.SetEmployeeDepartments Departments
    
    'Assert:
    Assert.IsTrue Employees.Count = 1
    Assert.IsTrue Employees.Item(1).DepartmentCount = 0

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Uncategorized")
Private Sub SetEmployeeDepartments_OneEmployeesOneDepartments_NoMatch()
    On Error GoTo TestFail
    
    'Arrange:
    Dim Employees As EmployeeCollection
    Dim Departments As DepartmentCollection
    Dim Emp As Employee
    Dim Dept As Department
    
    Set Employees = New EmployeeCollection
    Set Departments = New DepartmentCollection
    Set Emp = New Employee
    Set Dept = New Department
    
    Emp.EmplID = "1"
    Employees.Add Emp
    
    Dept.DeptID = "1"
    Dept.ManagerID = "2"
    
    'Act:
    Employees.SetEmployeeDepartments Departments
    
    'Assert:
    Assert.IsTrue Employees.Count = 1
    Assert.IsTrue Employees.Item(1).DepartmentCount = 0

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Uncategorized")
Private Sub SetEmployeeDepartments_OneEmployeesOneDepartments()
    On Error GoTo TestFail
    
    'Arrange:
    Dim Employees As EmployeeCollection
    Dim Departments As DepartmentCollection
    Dim Emp As Employee
    Dim Dept As Department
    
    Set Employees = New EmployeeCollection
    Set Departments = New DepartmentCollection
    Set Emp = New Employee
    Set Dept = New Department
    
    Emp.EmplID = "1"
    Employees.Add Emp
    
    Dept.DeptID = "1"
    Dept.ManagerID = "1"
    
    'Act:
    Employees.SetEmployeeDepartments Departments
    
    'Assert:
    Assert.IsTrue Employees.Count = 1
    Assert.IsTrue Employees.Item(1).DepartmentCount = 0

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Uncategorized")
Private Sub SetEmployeeDepartments_ExistingDepartments()
    On Error GoTo TestFail
    
    'Arrange:
    Dim Employees As EmployeeCollection
    Set Employees = New EmployeeCollection
    
    Dim Departments As DepartmentCollection
    Set Departments = New DepartmentCollection
    
    Dim Emp As Employee
    Set Emp = New Employee
    
    Dim Dept_Original As Department
    Dim Dept_Replacement As Department
    Set Dept_Original = New Department
    Set Dept_Replacement = New Department
    
    Emp.EmplID = "1"
    Dept_Original.ManagerID = "1"
    Dept_Replacement.ManagerID = "1"
    
    Dept_Original.DeptID = "0"
    Dept_Replacement.DeptID = "2"
    
    Emp.AddDepartment Dept_Original
    Employees.Add Emp
    
    Departments.Add Dept_Replacement
    
    'Act:
    Employees.SetEmployeeDepartments Departments
    
    'Assert:
    With Employees.Item(1)
        Assert.IsTrue .DepartmentCount = 1
        Assert.IsTrue .Department(1).DeptID = "2"
    End With

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub
