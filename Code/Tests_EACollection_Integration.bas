Attribute VB_Name = "TestModule_EAC_DC"
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
'@IgnoreModule UseMeaningfulName

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

'@TestMethod("Integration")
Private Sub TestMethod_OneMismatch_CorrectCount()
    On Error GoTo TestFail
    
    'Arrange:
    Dim DC As DepartmentCollection
    Dim DCMismatched As DepartmentCollection
    Dim EAC As ExpenseApprovalCollection
    Dim Dept As Department
    Dim EA As ExpenseApproval
    
    Set DC = New DepartmentCollection
    Set EAC = New ExpenseApprovalCollection
    Set Dept = New Department
    Set EA = New ExpenseApproval
    
    Dept.DeptID = "1"
    Dept.ManagerID = "2"
    EA.ApproverType = "EXAPPROVER"
    EA.FromChartfield = "1"
    EA.ToChartfield = "1"
    EA.EmplID = "1"
    
    DC.Add Dept
    EAC.Add EA
    
    'Act:
    Set DCMismatched = DC.DepartmentsWithExpenseApproverMismatch(EAC)
    
    'Assert:
    Assert.IsTrue 1 = DCMismatched.Count
    
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("No Fail")
Private Sub TestMethod_OneDepartment_NoExpenseApprovals_NoFail()
    On Error GoTo TestFail
    
    'Arrange:
    Dim ExpenseApprovers As ExpenseApprovalCollection
    Dim Departments As DepartmentCollection
    Dim Dept As Department
    
    Set ExpenseApprovers = New ExpenseApprovalCollection
    Set Departments = New DepartmentCollection
    Set Dept = New Department
    
    Dept.DeptID = "1"
    Dept.Description = "Test"
    Dept.ManagerID = "1"
    
    Departments.Add Dept
    
    'Act:
    '@Ignore FunctionReturnValueDiscarded
    Departments.DepartmentsWithExpenseApproverMismatch ExpenseApprovers
    
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
Private Sub TestMethod_OneDepartment_NoExpenseApprovals_ZeroCount()
    On Error GoTo TestFail
    
    'Arrange:
    Dim ExpenseApprovers As ExpenseApprovalCollection
    Dim Departments As DepartmentCollection
    Dim MismatchedDepartments As DepartmentCollection
    Dim Dept As Department
    
    Set ExpenseApprovers = New ExpenseApprovalCollection
    Set Departments = New DepartmentCollection
    Set Dept = New Department
    
    Dept.DeptID = "1"
    Dept.Description = "Test"
    Dept.ManagerID = "1"
    
    Departments.Add Dept
    
    'Act:
    Set MismatchedDepartments = Departments.DepartmentsWithExpenseApproverMismatch(ExpenseApprovers)
    
    'Assert:
    Assert.IsTrue MismatchedDepartments.Count = 0

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Uncategorized")
Private Sub TestMethod_NoDepartments_NoExpenseApprovers()
    On Error GoTo TestFail
    
    'Arrange:
    Dim ExpenseApprovers As ExpenseApprovalCollection
    Dim Departments As DepartmentCollection
    Dim MismatchedDepartments As DepartmentCollection
    
    Set ExpenseApprovers = New ExpenseApprovalCollection
    Set Departments = New DepartmentCollection
    
    'Act:
    Set MismatchedDepartments = Departments.DepartmentsWithExpenseApproverMismatch(ExpenseApprovers)
    
    'Assert:
    Assert.IsTrue MismatchedDepartments.Count = 0

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Uncategorized")
Private Sub TestMethod_NoDepartments_OneExpenseApprovers()
    On Error GoTo TestFail
    
    'Arrange:
    Dim ExpenseApprovers As ExpenseApprovalCollection
    Dim Departments As DepartmentCollection
    Dim MismatchedDepartments As DepartmentCollection
    Dim ExAppr As ExpenseApproval
    
    Set ExpenseApprovers = New ExpenseApprovalCollection
    Set Departments = New DepartmentCollection
    Set ExAppr = New ExpenseApproval
    
    ExAppr.ApproverType = "EXAPPROVER"
    ExAppr.EmplID = "1"
    ExAppr.FromChartfield = "1"
    ExAppr.ToChartfield = "2"
    
    'Act:
    Set MismatchedDepartments = Departments.DepartmentsWithExpenseApproverMismatch(ExpenseApprovers)
    
    'Assert:
    Assert.IsTrue MismatchedDepartments.Count = 0

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Uncategorized")
Private Sub TestMethod_TenDepts_TenExpenseApprovers_NoMismatch()
    On Error GoTo TestFail
    
    'Arrange:
    Dim ExpenseApprovers As ExpenseApprovalCollection
    Dim Departments As DepartmentCollection
    Dim MismatchedDepartments As DepartmentCollection
    Dim D0 As Department
    Dim D1 As Department
    Dim D2 As Department
    Dim D3 As Department
    Dim D4 As Department
    Dim D5 As Department
    Dim D6 As Department
    Dim D7 As Department
    Dim D8 As Department
    Dim D9 As Department
    Dim EA0 As ExpenseApproval
    Dim EA1 As ExpenseApproval
    Dim EA2 As ExpenseApproval
    Dim EA3 As ExpenseApproval
    Dim EA4 As ExpenseApproval
    Dim EA5 As ExpenseApproval
    Dim EA6 As ExpenseApproval
    Dim EA7 As ExpenseApproval
    Dim EA8 As ExpenseApproval
    Dim EA9 As ExpenseApproval
    
    Set ExpenseApprovers = New ExpenseApprovalCollection
    Set Departments = New DepartmentCollection
    
    Set D0 = New Department
    Set D1 = New Department
    Set D2 = New Department
    Set D3 = New Department
    Set D4 = New Department
    Set D5 = New Department
    Set D6 = New Department
    Set D7 = New Department
    Set D8 = New Department
    Set D9 = New Department
    Set EA0 = New ExpenseApproval
    Set EA1 = New ExpenseApproval
    Set EA2 = New ExpenseApproval
    Set EA3 = New ExpenseApproval
    Set EA4 = New ExpenseApproval
    Set EA5 = New ExpenseApproval
    Set EA6 = New ExpenseApproval
    Set EA7 = New ExpenseApproval
    Set EA8 = New ExpenseApproval
    Set EA9 = New ExpenseApproval
    
    D0.DeptID = "0"
    D1.DeptID = "1"
    D2.DeptID = "2"
    D3.DeptID = "3"
    D4.DeptID = "4"
    D5.DeptID = "5"
    D6.DeptID = "6"
    D7.DeptID = "7"
    D8.DeptID = "8"
    D9.DeptID = "9"
    D0.ManagerID = "0"
    D1.ManagerID = "1"
    D2.ManagerID = "2"
    D3.ManagerID = "3"
    D4.ManagerID = "4"
    D5.ManagerID = "5"
    D6.ManagerID = "6"
    D7.ManagerID = "7"
    D8.ManagerID = "8"
    D9.ManagerID = "9"
    
    EA0.ApproverType = "EXAPPROVER"
    EA1.ApproverType = "EXAPPROVER"
    EA2.ApproverType = "EXAPPROVER"
    EA3.ApproverType = "EXAPPROVER"
    EA4.ApproverType = "EXAPPROVER"
    EA5.ApproverType = "EXAPPROVER"
    EA6.ApproverType = "EXAPPROVER"
    EA7.ApproverType = "EXAPPROVER"
    EA8.ApproverType = "EXAPPROVER"
    EA9.ApproverType = "EXAPPROVER"
    
    EA0.FromChartfield = "0"
    EA0.ToChartfield = "0"
    EA0.EmplID = "0"
    EA1.FromChartfield = "1"
    EA1.ToChartfield = "1"
    EA1.EmplID = "1"
    EA2.FromChartfield = "2"
    EA2.ToChartfield = "2"
    EA2.EmplID = "2"
    EA3.FromChartfield = "3"
    EA3.ToChartfield = "3"
    EA3.EmplID = "3"
    EA4.FromChartfield = "4"
    EA4.ToChartfield = "4"
    EA4.EmplID = "4"
    EA5.FromChartfield = "5"
    EA5.ToChartfield = "5"
    EA5.EmplID = "5"
    EA6.FromChartfield = "6"
    EA6.ToChartfield = "6"
    EA6.EmplID = "6"
    EA7.FromChartfield = "7"
    EA7.ToChartfield = "7"
    EA7.EmplID = "7"
    EA8.FromChartfield = "8"
    EA8.ToChartfield = "8"
    EA8.EmplID = "8"
    EA9.FromChartfield = "9"
    EA9.ToChartfield = "9"
    EA9.EmplID = "9"
    
    Departments.Add D0
    Departments.Add D1
    Departments.Add D2
    Departments.Add D3
    Departments.Add D4
    Departments.Add D5
    Departments.Add D6
    Departments.Add D7
    Departments.Add D8
    Departments.Add D9
    
    ExpenseApprovers.Add EA0
    ExpenseApprovers.Add EA1
    ExpenseApprovers.Add EA2
    ExpenseApprovers.Add EA3
    ExpenseApprovers.Add EA4
    ExpenseApprovers.Add EA5
    ExpenseApprovers.Add EA6
    ExpenseApprovers.Add EA7
    ExpenseApprovers.Add EA8
    ExpenseApprovers.Add EA9
    
    'Act:
    Set MismatchedDepartments = Departments.DepartmentsWithExpenseApproverMismatch(ExpenseApprovers)
    
    'Assert:
    Assert.IsTrue MismatchedDepartments.Count = 0

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("No Fail")
Private Sub TestMethod_TenDepts_TenExpenseApprovers_NoFail()
    On Error GoTo TestFail
    
    'Arrange:
    Dim ExpenseApprovers As ExpenseApprovalCollection
    Dim Departments As DepartmentCollection
    '@Ignore VariableNotUsed
    Dim MismatchedDepartments As DepartmentCollection
    Dim D0 As Department
    Dim D1 As Department
    Dim D2 As Department
    Dim D3 As Department
    Dim D4 As Department
    Dim D5 As Department
    Dim D6 As Department
    Dim D7 As Department
    Dim D8 As Department
    Dim D9 As Department
    Dim EA0 As ExpenseApproval
    Dim EA1 As ExpenseApproval
    Dim EA2 As ExpenseApproval
    Dim EA3 As ExpenseApproval
    Dim EA4 As ExpenseApproval
    Dim EA5 As ExpenseApproval
    Dim EA6 As ExpenseApproval
    Dim EA7 As ExpenseApproval
    Dim EA8 As ExpenseApproval
    Dim EA9 As ExpenseApproval
    
    Set ExpenseApprovers = New ExpenseApprovalCollection
    Set Departments = New DepartmentCollection
    
    Set D0 = New Department
    Set D1 = New Department
    Set D2 = New Department
    Set D3 = New Department
    Set D4 = New Department
    Set D5 = New Department
    Set D6 = New Department
    Set D7 = New Department
    Set D8 = New Department
    Set D9 = New Department
    Set EA0 = New ExpenseApproval
    Set EA1 = New ExpenseApproval
    Set EA2 = New ExpenseApproval
    Set EA3 = New ExpenseApproval
    Set EA4 = New ExpenseApproval
    Set EA5 = New ExpenseApproval
    Set EA6 = New ExpenseApproval
    Set EA7 = New ExpenseApproval
    Set EA8 = New ExpenseApproval
    Set EA9 = New ExpenseApproval
    
    D0.DeptID = "0"
    D1.DeptID = "1"
    D2.DeptID = "2"
    D3.DeptID = "3"
    D4.DeptID = "4"
    D5.DeptID = "5"
    D6.DeptID = "6"
    D7.DeptID = "7"
    D8.DeptID = "8"
    D9.DeptID = "9"
    D0.ManagerID = "0"
    D1.ManagerID = "1"
    D2.ManagerID = "2"
    D3.ManagerID = "3"
    D4.ManagerID = "4"
    D5.ManagerID = "5"
    D6.ManagerID = "6"
    D7.ManagerID = "7"
    D8.ManagerID = "8"
    D9.ManagerID = "9"
    
    EA0.ApproverType = "EXAPPROVER"
    EA1.ApproverType = "EXAPPROVER"
    EA2.ApproverType = "EXAPPROVER"
    EA3.ApproverType = "EXAPPROVER"
    EA4.ApproverType = "EXAPPROVER"
    EA5.ApproverType = "EXAPPROVER"
    EA6.ApproverType = "EXAPPROVER"
    EA7.ApproverType = "EXAPPROVER"
    EA8.ApproverType = "EXAPPROVER"
    EA9.ApproverType = "EXAPPROVER"
    
    EA0.FromChartfield = "0"
    EA0.ToChartfield = "0"
    EA0.EmplID = "0"
    EA1.FromChartfield = "1"
    EA1.ToChartfield = "1"
    EA1.EmplID = "1"
    EA2.FromChartfield = "2"
    EA2.ToChartfield = "2"
    EA2.EmplID = "2"
    EA3.FromChartfield = "3"
    EA3.ToChartfield = "3"
    EA3.EmplID = "3"
    EA4.FromChartfield = "4"
    EA4.ToChartfield = "4"
    EA4.EmplID = "4"
    EA5.FromChartfield = "5"
    EA5.ToChartfield = "5"
    EA5.EmplID = "5"
    EA6.FromChartfield = "6"
    EA6.ToChartfield = "6"
    EA6.EmplID = "6"
    EA7.FromChartfield = "7"
    EA7.ToChartfield = "7"
    EA7.EmplID = "7"
    EA8.FromChartfield = "8"
    EA8.ToChartfield = "8"
    EA8.EmplID = "8"
    EA9.FromChartfield = "9"
    EA9.ToChartfield = "9"
    EA9.EmplID = "9"
    
    Departments.Add D0
    Departments.Add D1
    Departments.Add D2
    Departments.Add D3
    Departments.Add D4
    Departments.Add D5
    Departments.Add D6
    Departments.Add D7
    Departments.Add D8
    Departments.Add D9
    
    ExpenseApprovers.Add EA0
    ExpenseApprovers.Add EA1
    ExpenseApprovers.Add EA2
    ExpenseApprovers.Add EA3
    ExpenseApprovers.Add EA4
    ExpenseApprovers.Add EA5
    ExpenseApprovers.Add EA6
    ExpenseApprovers.Add EA7
    ExpenseApprovers.Add EA8
    ExpenseApprovers.Add EA9
    
    'Act:
    '@Ignore AssignmentNotUsed
    Set MismatchedDepartments = Departments.DepartmentsWithExpenseApproverMismatch(ExpenseApprovers)
    
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
Private Sub TestMethod_TenDepts_TenExpenseApprovers_OneMismatch()
    On Error GoTo TestFail
    
    'Arrange:
    Dim ExpenseApprovers As ExpenseApprovalCollection
    Dim Departments As DepartmentCollection
    Dim MismatchedDepartments As DepartmentCollection
    Dim D0 As Department
    Dim D1 As Department
    Dim D2 As Department
    Dim D3 As Department
    Dim D4 As Department
    Dim D5 As Department
    Dim D6 As Department
    Dim D7 As Department
    Dim D8 As Department
    Dim D9 As Department
    Dim EA0 As ExpenseApproval
    Dim EA1 As ExpenseApproval
    Dim EA2 As ExpenseApproval
    Dim EA3 As ExpenseApproval
    Dim EA4 As ExpenseApproval
    Dim EA5 As ExpenseApproval
    Dim EA6 As ExpenseApproval
    Dim EA7 As ExpenseApproval
    Dim EA8 As ExpenseApproval
    Dim EA9 As ExpenseApproval
    
    Set ExpenseApprovers = New ExpenseApprovalCollection
    Set Departments = New DepartmentCollection
    
    Set D0 = New Department
    Set D1 = New Department
    Set D2 = New Department
    Set D3 = New Department
    Set D4 = New Department
    Set D5 = New Department
    Set D6 = New Department
    Set D7 = New Department
    Set D8 = New Department
    Set D9 = New Department
    Set EA0 = New ExpenseApproval
    Set EA1 = New ExpenseApproval
    Set EA2 = New ExpenseApproval
    Set EA3 = New ExpenseApproval
    Set EA4 = New ExpenseApproval
    Set EA5 = New ExpenseApproval
    Set EA6 = New ExpenseApproval
    Set EA7 = New ExpenseApproval
    Set EA8 = New ExpenseApproval
    Set EA9 = New ExpenseApproval
    
    D0.DeptID = "0"
    D1.DeptID = "1"
    D2.DeptID = "2"
    D3.DeptID = "3"
    D4.DeptID = "4"
    D5.DeptID = "5"
    D6.DeptID = "6"
    D7.DeptID = "7"
    D8.DeptID = "8"
    D9.DeptID = "9"
    D0.ManagerID = "0"
    D1.ManagerID = "1"
    D2.ManagerID = "2"
    D3.ManagerID = "3"
    D4.ManagerID = "4"
    D5.ManagerID = "5"
    D6.ManagerID = "6"
    D7.ManagerID = "7"
    D8.ManagerID = "8"
    D9.ManagerID = "99"
    
    EA0.ApproverType = "EXAPPROVER"
    EA1.ApproverType = "EXAPPROVER"
    EA2.ApproverType = "EXAPPROVER"
    EA3.ApproverType = "EXAPPROVER"
    EA4.ApproverType = "EXAPPROVER"
    EA5.ApproverType = "EXAPPROVER"
    EA6.ApproverType = "EXAPPROVER"
    EA7.ApproverType = "EXAPPROVER"
    EA8.ApproverType = "EXAPPROVER"
    EA9.ApproverType = "EXAPPROVER"
    
    EA0.FromChartfield = "0"
    EA0.ToChartfield = "0"
    EA0.EmplID = "0"
    EA1.FromChartfield = "1"
    EA1.ToChartfield = "1"
    EA1.EmplID = "1"
    EA2.FromChartfield = "2"
    EA2.ToChartfield = "2"
    EA2.EmplID = "2"
    EA3.FromChartfield = "3"
    EA3.ToChartfield = "3"
    EA3.EmplID = "3"
    EA4.FromChartfield = "4"
    EA4.ToChartfield = "4"
    EA4.EmplID = "4"
    EA5.FromChartfield = "5"
    EA5.ToChartfield = "5"
    EA5.EmplID = "5"
    EA6.FromChartfield = "6"
    EA6.ToChartfield = "6"
    EA6.EmplID = "6"
    EA7.FromChartfield = "7"
    EA7.ToChartfield = "7"
    EA7.EmplID = "7"
    EA8.FromChartfield = "8"
    EA8.ToChartfield = "8"
    EA8.EmplID = "8"
    EA9.FromChartfield = "9"
    EA9.ToChartfield = "9"
    EA9.EmplID = "9"
    
    Departments.Add D0
    Departments.Add D1
    Departments.Add D2
    Departments.Add D3
    Departments.Add D4
    Departments.Add D5
    Departments.Add D6
    Departments.Add D7
    Departments.Add D8
    Departments.Add D9
    
    ExpenseApprovers.Add EA0
    ExpenseApprovers.Add EA1
    ExpenseApprovers.Add EA2
    ExpenseApprovers.Add EA3
    ExpenseApprovers.Add EA4
    ExpenseApprovers.Add EA5
    ExpenseApprovers.Add EA6
    ExpenseApprovers.Add EA7
    ExpenseApprovers.Add EA8
    ExpenseApprovers.Add EA9
    
    'Act:
    Set MismatchedDepartments = Departments.DepartmentsWithExpenseApproverMismatch(ExpenseApprovers)
    
    'Assert:
    Assert.IsTrue MismatchedDepartments.Count = 1

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Uncategorized")
Private Sub TestMethod_DepartmentsWithoutExpenseApproval_NoMatches_OneDeptOneEA()
    On Error GoTo TestFail
    
    'Arrange:
    Dim Departments As DepartmentCollection
    Dim ExApprovals As ExpenseApprovalCollection
    
    Dim Dept As Department
    Dim ExAppr As ExpenseApproval
    
    Set Departments = New DepartmentCollection
    Set ExApprovals = New ExpenseApprovalCollection
    Set Dept = New Department
    Set ExAppr = New ExpenseApproval
    
    Dept.DeptID = "1"
    ExAppr.ApproverType = "EXAPPROVER"
    ExAppr.FromChartfield = "1"
    ExAppr.ToChartfield = "1"
    Departments.Add Dept
    ExApprovals.Add ExAppr
    
    'Act:
    'Assert:
    Assert.IsTrue Departments.DepartmentsWithoutExpenseApproval(ExApprovals).Count = 0

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Uncategorized")
Private Sub TestMethod_DepartmentsWithoutExpenseApproval_OneMatches_OneDeptOneEA_CountCorrect()
    On Error GoTo TestFail
    
    'Arrange:
    Dim Departments As DepartmentCollection
    Dim ExApprovals As ExpenseApprovalCollection
    
    Dim Dept As Department
    Dim ExAppr As ExpenseApproval
    
    Set Departments = New DepartmentCollection
    Set ExApprovals = New ExpenseApprovalCollection
    Set Dept = New Department
    Set ExAppr = New ExpenseApproval
    
    Dept.DeptID = "1"
    ExAppr.ApproverType = "EXAPPROVER"
    ExAppr.FromChartfield = "2"
    ExAppr.ToChartfield = "2"
    Departments.Add Dept
    ExApprovals.Add ExAppr
    
    'Act:
    'Assert:
    Assert.IsTrue Departments.DepartmentsWithoutExpenseApproval(ExApprovals).Count = 1

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Uncategorized")
Private Sub TestMethod_DepartmentsWithoutExpenseApproval_OneMatches_OneDeptOneEA_DeptIDCorrect()
    On Error GoTo TestFail
    
    'Arrange:
    Dim Departments As DepartmentCollection
    Dim ExApprovals As ExpenseApprovalCollection
    
    Dim Dept As Department
    Dim ExAppr As ExpenseApproval
    
    Set Departments = New DepartmentCollection
    Set ExApprovals = New ExpenseApprovalCollection
    Set Dept = New Department
    Set ExAppr = New ExpenseApproval
    
    Dept.DeptID = "1"
    ExAppr.ApproverType = "EXAPPROVER"
    ExAppr.FromChartfield = "2"
    ExAppr.ToChartfield = "2"
    Departments.Add Dept
    ExApprovals.Add ExAppr
    
    'Act:
    'Assert:
    Assert.IsTrue Departments.DepartmentsWithoutExpenseApproval(ExApprovals).Item(1).DeptID = "1"

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Integration")
'@Ignore UseMeaningfulName
Private Sub TestMethod_DepartmentHasEXApprover_OneEAInCollection_Range1()
    On Error GoTo TestFail
    
    'Arrange:
    Dim ExpenseApprovers As ExpenseApprovalCollection
    '@Ignore UseMeaningfulName
    Dim EA1 As ExpenseApproval
    Dim Dept As Department
    Set ExpenseApprovers = New ExpenseApprovalCollection
    Set EA1 = New ExpenseApproval
    Set Dept = New Department
    
    EA1.ApproverType = "EXAPPROVER"
    EA1.FromChartfield = "1"
    EA1.ToChartfield = "1"
    Dept.DeptID = "1"
    
    ExpenseApprovers.Add EA1
    
    'Act:
    'Assert:
    Assert.IsTrue Dept.DepartmentHasEXApproverInCollection(ExpenseApprovers)

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Integration")
Private Sub TestMethod_DepartmentHasEXApprover_OneEAInCollection_FromChartfield()
    On Error GoTo TestFail
    
    'Arrange:
    Dim ExpenseApprovers As ExpenseApprovalCollection
    '@Ignore UseMeaningfulName
    Dim EA1 As ExpenseApproval
    Dim Dept As Department
    Set ExpenseApprovers = New ExpenseApprovalCollection
    Set EA1 = New ExpenseApproval
    Set Dept = New Department
    
    EA1.ApproverType = "EXAPPROVER"
    EA1.FromChartfield = "1"
    EA1.ToChartfield = "2"
    Dept.DeptID = "1"
    
    ExpenseApprovers.Add EA1
    
    'Act:
    'Assert:
    Assert.IsTrue Dept.DepartmentHasEXApproverInCollection(ExpenseApprovers)

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Integration")
Private Sub TestMethod_DepartmentHasEXApprover_OneEAInCollection_ToChartfield()
    On Error GoTo TestFail
    
    'Arrange:
    Dim ExpenseApprovers As ExpenseApprovalCollection
    '@Ignore UseMeaningfulName
    Dim EA1 As ExpenseApproval
    Dim Dept As Department
    Set ExpenseApprovers = New ExpenseApprovalCollection
    Set EA1 = New ExpenseApproval
    Set Dept = New Department
    
    EA1.ApproverType = "EXAPPROVER"
    EA1.FromChartfield = "1"
    EA1.ToChartfield = "2"
    Dept.DeptID = "2"
    
    ExpenseApprovers.Add EA1
    
    'Act:
    'Assert:
    Assert.IsTrue Dept.DepartmentHasEXApproverInCollection(ExpenseApprovers)

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("Integration")
'@Ignore UseMeaningfulName
Private Sub TestMethod_DepartmentHasVPApprover_OneEA_Range1()
    On Error GoTo TestFail
    
    'Arrange:
    Dim ExpenseApprovers As ExpenseApprovalCollection
    '@Ignore UseMeaningfulName
    Dim EA1 As ExpenseApproval
    Dim Dept As Department
    Set ExpenseApprovers = New ExpenseApprovalCollection
    Set EA1 = New ExpenseApproval
    Set Dept = New Department
    
    EA1.ApproverType = "VPAPPROVER"
    EA1.FromChartfield = "1"
    EA1.ToChartfield = "1"
    Dept.DeptID = "1"
    
    ExpenseApprovers.Add EA1
    
    'Act:
    'Assert:
    Assert.IsTrue Dept.DepartmentHasVPApproverInCollection(ExpenseApprovers)

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub
