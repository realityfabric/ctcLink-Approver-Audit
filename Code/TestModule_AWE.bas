Attribute VB_Name = "TestModule_AWE"
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
'@Folder("Tests.IO")

Option Explicit
Option Private Module

Private Assert As Object
Private wbAWE As Workbook

'@ModuleInitialize
Private Sub ModuleInitialize()
    'this method runs once per module.
    Set Assert = CreateObject("Rubberduck.AssertClass")
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
    'this method runs once per module.
    Set Assert = Nothing
    Set wbAWE = Nothing
End Sub

'@TestInitialize
Private Sub TestInitialize()
    Set wbAWE = Workbooks.Open(ThisWorkbook.Path & "/test_data/QFS_SEC_EOAW_APPROVAL_SETUP.csv", ReadOnly:=True)
End Sub

'@TestCleanup
Private Sub TestCleanup()
    Dim DisplayAlerts As Boolean
    DisplayAlerts = Application.DisplayAlerts
    Application.DisplayAlerts = False
    wbAWE.Close SaveChanges:=False
    Application.DisplayAlerts = DisplayAlerts
End Sub

'@TestMethod("No Fail")
Private Sub TestMethod_ReadFromQuery_NoFail()
    On Error GoTo TestFail
    
    'Arrange:
    Dim AWE As ApprovalWorkflowConfig
    Dim Depts As DepartmentCollection
    Set AWE = New ApprovalWorkflowConfig
    Set Depts = New DepartmentCollection
    
    'Act:
    AWE.ReadFrom_QFS_SEC_EOAW_APPROVAL_SETUP_sheet wbAWE.Sheets.Item(1), 3, Depts
    
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
Private Sub TestMethod_ReadFromQuery_NoDepts_CountZero()
    On Error GoTo TestFail
    
    'Arrange:
    Dim AWE As ApprovalWorkflowConfig
    Dim Depts As DepartmentCollection
    Set AWE = New ApprovalWorkflowConfig
    Set Depts = New DepartmentCollection
    
    'Act:
    AWE.ReadFrom_QFS_SEC_EOAW_APPROVAL_SETUP_sheet wbAWE.Sheets.Item(1), 3, Depts
    
    'Assert:
    Assert.IsTrue AWE.ValueCount = 0

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next

    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Uncategorized")
Private Sub TestMethod_ReadFromQuery_NoDepts_ProcessIDCorrect()
    On Error GoTo TestFail
    
    'Arrange:
    Dim AWE As ApprovalWorkflowConfig
    Dim Depts As DepartmentCollection
    Set AWE = New ApprovalWorkflowConfig
    Set Depts = New DepartmentCollection
    
    'Act:
    AWE.ReadFrom_QFS_SEC_EOAW_APPROVAL_SETUP_sheet wbAWE.Sheets.Item(1), 3, Depts
    
    'Assert:
    Assert.IsTrue "PurchaseOrder" = AWE.ProcessID
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next

    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Uncategorized")
Private Sub TestMethod_ReadFromQuery_NoDepts_DefinitionIDCorrect()
    On Error GoTo TestFail
    
    'Arrange:
    Dim AWE As ApprovalWorkflowConfig
    Dim Depts As DepartmentCollection
    Set AWE = New ApprovalWorkflowConfig
    Set Depts = New DepartmentCollection
    
    'Act:
    AWE.ReadFrom_QFS_SEC_EOAW_APPROVAL_SETUP_sheet wbAWE.Sheets.Item(1), 3, Depts
    
    'Assert:
    Assert.IsTrue "SHARE" = AWE.DefinitionID
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next

    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub
