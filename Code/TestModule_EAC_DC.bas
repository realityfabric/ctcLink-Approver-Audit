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

Option Explicit
Option Private Module

Private Assert As Object

'@ModuleInitialize
Private Sub ModuleInitialize()
    'this method runs once per module.
    Set Assert = CreateObject("Rubberduck.AssertClass")
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
    'this method runs once per module.
    Set Assert = Nothing
End Sub

'@TestMethod("Integration")
Private Sub TestMethod_OneMismatch_CorrectCount()
    On Error GoTo TestFail
    
    'Arrange:
    '@Ignore UseMeaningfulName
    Dim DC As DepartmentCollection
    Dim DCMismatched As DepartmentCollection
    '@Ignore UseMeaningfulName
    Dim EAC As ExpenseApprovalCollection
    Dim Dept As Department
    '@Ignore UseMeaningfulName
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
