Attribute VB_Name = "Tests_EA_Integration"
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
