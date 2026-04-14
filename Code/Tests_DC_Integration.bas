Attribute VB_Name = "Tests_DC_Integration"
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

'@TestMethod("No Fail")
Private Sub DepartmentsWithoutVPApproval_NoDepartments_NoConfigs_NoFail()
    On Error GoTo TestFail
    
    'Arrange:
    Dim Departments As DepartmentCollection
    Dim Configs As AWConfigCollection
    Set Departments = New DepartmentCollection
    Set Configs = New AWConfigCollection
    
    'Act:
    '@Ignore FunctionReturnValueDiscarded
    Departments.DepartmentsWithoutVPApproval Configs
    
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
Private Sub DepartmentsWithoutVPApproval_NoDepartments_NoFail()
    On Error GoTo TestFail
    
    'Arrange:
    Dim Departments As DepartmentCollection
    Dim Configs As AWConfigCollection
    Dim Conf(1 To 5) As ApprovalWorkflowConfig
    Dim Index As Long
    Set Departments = New DepartmentCollection
    Set Configs = New AWConfigCollection
    
    For Index = 1 To 5
        Set Conf(Index) = New ApprovalWorkflowConfig
        With Conf(Index)
            .DefinitionID = "WA999"
            .Approvers = "AW_PO_EXEC_LEVEL_99"
            .Description = "Test #" & Str$(Index)
            .EffectiveDate = #1/1/2026#
            .EffectiveStatus = "A"
            .ProcessID = "VoucherApproval"
            .Stage = 10
            .Path = 1
            .Step = 1
            .StepFieldName = "DEPTID"
            .AddValue Trim$(Str$(Index))
        End With
        
        Configs.Add Conf(Index)
    Next Index
    
    'Act:
    '@Ignore FunctionReturnValueDiscarded
    Departments.DepartmentsWithoutVPApproval Configs
    
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
Private Sub DepartmentsWithoutVPApproval_NoDepartments_CountCorrect()
    On Error GoTo TestFail
    
    'Arrange:
    Dim Departments As DepartmentCollection
    Dim Configs As AWConfigCollection
    Dim Conf(1 To 5) As ApprovalWorkflowConfig
    Dim Index As Long
    Set Departments = New DepartmentCollection
    Set Configs = New AWConfigCollection
    
    For Index = 1 To 5
        Set Conf(Index) = New ApprovalWorkflowConfig
        With Conf(Index)
            .DefinitionID = "WA999"
            .Approvers = "AW_PO_EXEC_LEVEL_99"
            .Description = "Test #" & Str$(Index)
            .EffectiveDate = #1/1/2026#
            .EffectiveStatus = "A"
            .ProcessID = "VoucherApproval"
            .Stage = 10
            .Path = 1
            .Step = 1
            .StepFieldName = "DEPTID"
            .AddValue Trim$(Str$(Index))
        End With
        
        Configs.Add Conf(Index)
    Next Index
    
    'Act:
    'Assert:
    Assert.IsTrue 0 = Departments.DepartmentsWithoutVPApproval(Configs).Count

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Uncategorized")
'@Ignore UseMeaningfulName
Private Sub DepartmentsWithoutVPApproval_OneDepartments_CountCorrect0()
    On Error GoTo TestFail
    
    'Arrange:
    Dim Departments As DepartmentCollection
    Dim Configs As AWConfigCollection
    Dim Conf(1 To 5) As ApprovalWorkflowConfig
    Dim Dept As Department
    Dim Index As Long
    Set Departments = New DepartmentCollection
    Set Configs = New AWConfigCollection
    Set Dept = New Department
    
    For Index = 1 To 5
        Set Conf(Index) = New ApprovalWorkflowConfig
        With Conf(Index)
            .DefinitionID = "WA999"
            .Approvers = "AW_PO_EXEC_LEVEL_99"
            .Description = "Test #" & Str$(Index)
            .EffectiveDate = #1/1/2026#
            .EffectiveStatus = "A"
            .ProcessID = "VoucherApproval"
            .Stage = 10
            .Path = 1
            .Step = 1
            .StepFieldName = "DEPTID"
            .AddValue Trim$(Str$(Index))
        End With
        
        Configs.Add Conf(Index)
    Next Index
    
    Dept.DeptID = "1"
    Dept.Description = "Test 1"
    Dept.ManagerID = "1"
    
    Departments.Add Dept
    
    'Act:
    'Assert:
    Assert.IsTrue 0 = Departments.DepartmentsWithoutVPApproval(Configs).Count

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Uncategorized")
'@Ignore UseMeaningfulName
Private Sub DepartmentsWithoutVPApproval_OneDepartments_CountCorrect1()
    On Error GoTo TestFail
    
    'Arrange:
    Dim Departments As DepartmentCollection
    Dim Configs As AWConfigCollection
    Dim Conf(1 To 5) As ApprovalWorkflowConfig
    Dim Dept As Department
    Dim Index As Long
    Set Departments = New DepartmentCollection
    Set Configs = New AWConfigCollection
    Set Dept = New Department
    
    For Index = 1 To 5
        Set Conf(Index) = New ApprovalWorkflowConfig
        With Conf(Index)
            .DefinitionID = "WA999"
            .Approvers = "AW_PO_EXEC_LEVEL_99"
            .Description = "Test #" & Str$(Index)
            .EffectiveDate = #1/1/2026#
            .EffectiveStatus = "A"
            .ProcessID = "VoucherApproval"
            .Stage = 10
            .Path = 1
            .Step = 1
            .StepFieldName = "DEPTID"
            .AddValue Trim$(Str$(Index))
        End With
        
        Configs.Add Conf(Index)
    Next Index
    
    Dept.DeptID = "1000"
    Dept.Description = "Test 1"
    Dept.ManagerID = "1"
    
    Departments.Add Dept
    
    'Act:
    'Assert:
    Assert.IsTrue 1 = Departments.DepartmentsWithoutVPApproval(Configs).Count

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub
