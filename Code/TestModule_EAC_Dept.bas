Attribute VB_Name = "TestModule_EAC_Dept"
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

'@TestMethod("Integration")
Private Sub TestMethod_DepartmentHasEXApprover_OneEAInCollection_Range1()
    On Error GoTo TestFail
    
    'Arrange:
    Dim ExpenseApprovers As ExpenseApprovalCollection
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
Private Sub TestMethod_DepartmentHasVPApprover_OneEA_Range1()
    On Error GoTo TestFail
    
    'Arrange:
    Debug.Print "Arrange"
    Dim ExpenseApprovers As ExpenseApprovalCollection
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
    Debug.Print "Act"
    'Assert:
    Debug.Print "Assert"
    Assert.IsTrue Dept.DepartmentHasVPApproverInCollection(ExpenseApprovers)

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub
