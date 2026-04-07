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


'@TestMethod("Uncategorized")
Private Sub TestMethod_ReadFromQuery_NoDepts_StepFieldCorrect()
    On Error GoTo TestFail
    
    'Arrange:
    '@Ignore UseMeaningfulName
    Dim AWE_3 As ApprovalWorkflowConfig
    '@Ignore UseMeaningfulName
    Dim AWE_4 As ApprovalWorkflowConfig
    '@Ignore UseMeaningfulName
    Dim AWE_17 As ApprovalWorkflowConfig
    '@Ignore UseMeaningfulName
    Dim AWE_92 As ApprovalWorkflowConfig
    Dim Depts As DepartmentCollection
    Set AWE_3 = New ApprovalWorkflowConfig
    Set AWE_4 = New ApprovalWorkflowConfig
    Set AWE_17 = New ApprovalWorkflowConfig
    Set AWE_92 = New ApprovalWorkflowConfig
    Set Depts = New DepartmentCollection
    
    'Act:
    AWE_3.ReadFrom_QFS_SEC_EOAW_APPROVAL_SETUP_sheet wbAWE.Sheets.Item(1), 3, Depts
    AWE_4.ReadFrom_QFS_SEC_EOAW_APPROVAL_SETUP_sheet wbAWE.Sheets.Item(1), 4, Depts
    AWE_17.ReadFrom_QFS_SEC_EOAW_APPROVAL_SETUP_sheet wbAWE.Sheets.Item(1), 17, Depts
    AWE_92.ReadFrom_QFS_SEC_EOAW_APPROVAL_SETUP_sheet wbAWE.Sheets.Item(1), 92, Depts
    
    'Assert:
    Assert.IsTrue vbNullString = AWE_3.StepFieldName
    Assert.IsTrue "CATEGORY_ID" = AWE_4.StepFieldName
    Assert.IsTrue "DEPTID" = AWE_17.StepFieldName
    Assert.IsTrue "FUND_CODE" = AWE_92.StepFieldName
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next

    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Uncategorized")
Private Sub TestMethod_ReadFromQuery_ThreeDepts_OperatorBetween_Count2Correct()
    On Error GoTo TestFail
    
    'Arrange:
    '@Ignore UseMeaningfulName
    Dim AWE_17 As ApprovalWorkflowConfig
    '@Ignore UseMeaningfulName
    Dim AWE_92 As ApprovalWorkflowConfig
    Dim Depts As DepartmentCollection
    Dim Dept1 As Department
    Dim Dept2 As Department
    Dim Dept3 As Department
    Set AWE_17 = New ApprovalWorkflowConfig
    Set Depts = New DepartmentCollection
    Set Dept1 = New Department
    Set Dept2 = New Department
    Set Dept3 = New Department
    
    Dept1.DeptID = "10601"
    Dept2.DeptID = "10603"
    Dept3.DeptID = "10605"
    
    Depts.Add Dept1
    Depts.Add Dept2
    Depts.Add Dept3

    'Act:
    AWE_17.ReadFrom_QFS_SEC_EOAW_APPROVAL_SETUP_sheet wbAWE.Sheets.Item(1), 17, Depts
    
    'Assert:
    Assert.IsTrue 2 = AWE_17.ValueCount

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next

    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Uncategorized")
Private Sub TestMethod_ReadFromQuery_ThreeDepts_OperatorBetween_ValuesCorrect()
    On Error GoTo TestFail
    
    'Arrange:
    '@Ignore UseMeaningfulName
    Dim AWE_17 As ApprovalWorkflowConfig
    '@Ignore UseMeaningfulName
    Dim AWE_92 As ApprovalWorkflowConfig
    Dim Depts As DepartmentCollection
    Dim Dept1 As Department
    Dim Dept2 As Department
    Dim Dept3 As Department
    Set AWE_17 = New ApprovalWorkflowConfig
    Set Depts = New DepartmentCollection
    Set Dept1 = New Department
    Set Dept2 = New Department
    Set Dept3 = New Department
    
    Dept1.DeptID = "10601"
    Dept2.DeptID = "10603"
    Dept3.DeptID = "10605"
    
    Depts.Add Dept1
    Depts.Add Dept2
    Depts.Add Dept3

    'Act:
    AWE_17.ReadFrom_QFS_SEC_EOAW_APPROVAL_SETUP_sheet wbAWE.Sheets.Item(1), 17, Depts
    
    'Assert:
    Debug.Print "Value 1: " & AWE_17.Value(1)
    Debug.Print "Value 2: " & AWE_17.Value(2)
    Assert.IsTrue ("10601" = AWE_17.Value(1) And "10603" = AWE_17.Value(2)) Or _
                  ("10603" = AWE_17.Value(1) And "10601" = AWE_17.Value(2))
    

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next

    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub
