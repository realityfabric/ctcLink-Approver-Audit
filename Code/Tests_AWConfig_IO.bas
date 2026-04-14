Attribute VB_Name = "Tests_AWConfig"
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
    Set Assert = CreateObject("Rubberduck.AssertClass")
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
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
    Dim Depts As DepartmentCollection
    '@Ignore UseMeaningfulName
    Dim Dept1 As Department
    '@Ignore UseMeaningfulName
    Dim Dept2 As Department
    '@Ignore UseMeaningfulName
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
    Dim Depts As DepartmentCollection
    '@Ignore UseMeaningfulName
    Dim Dept1 As Department
    '@Ignore UseMeaningfulName
    Dim Dept2 As Department
    '@Ignore UseMeaningfulName
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

'@TestMethod("Uncategorized")
Private Sub TestMethod_ReadFromQuery_ThreeDepts_OperatorEquals_ValuesCorrect()
    On Error GoTo TestFail
    
    'Arrange:
    '@Ignore UseMeaningfulName
    Dim AWE_18 As ApprovalWorkflowConfig
    Dim Depts As DepartmentCollection
    '@Ignore UseMeaningfulName
    Dim Dept1 As Department
    '@Ignore UseMeaningfulName
    Dim Dept2 As Department
    '@Ignore UseMeaningfulName
    Dim Dept3 As Department
    Set AWE_18 = New ApprovalWorkflowConfig
    Set Depts = New DepartmentCollection
    Set Dept1 = New Department
    Set Dept2 = New Department
    Set Dept3 = New Department
    
    Dept1.DeptID = "10601"
    Dept2.DeptID = "60027"
    Dept3.DeptID = "10605"
    
    Depts.Add Dept1
    Depts.Add Dept2
    Depts.Add Dept3

    'Act:
    AWE_18.ReadFrom_QFS_SEC_EOAW_APPROVAL_SETUP_sheet wbAWE.Sheets.Item(1), 18, Depts
    
    'Assert:
    Assert.IsTrue AWE_18.Value(1) = "60027"
    

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next

    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("Uncategorized")
Private Sub TestMethod_ReadFromQuery_ThreeDepts_OperatorList_ValuesCorrect()
    On Error GoTo TestFail
    
    'Arrange:
    '@Ignore UseMeaningfulName
    Dim AWE_65 As ApprovalWorkflowConfig
    Dim Depts As DepartmentCollection
    '@Ignore UseMeaningfulName
    Dim Dept1 As Department
    '@Ignore UseMeaningfulName
    Dim Dept2 As Department
    '@Ignore UseMeaningfulName
    Dim Dept3 As Department
    Set AWE_65 = New ApprovalWorkflowConfig
    Set Depts = New DepartmentCollection
    Set Dept1 = New Department
    Set Dept2 = New Department
    Set Dept3 = New Department
    
    Dept1.DeptID = "20240"
    Dept2.DeptID = "50100"
    Dept3.DeptID = "10605"
    
    Depts.Add Dept1
    Depts.Add Dept2
    Depts.Add Dept3

    'Act:
    AWE_65.ReadFrom_QFS_SEC_EOAW_APPROVAL_SETUP_sheet wbAWE.Sheets.Item(1), 65, Depts
    
    'Assert:
    Debug.Print "|" & AWE_65.ValueCount & "|"
    Debug.Print "|" & AWE_65.Value(1) & "|"
    Debug.Print "|" & AWE_65.Value(2) & "|"
    Assert.IsTrue (AWE_65.Value(1) = "20240" And AWE_65.Value(2) = "50100") Or _
                  (AWE_65.Value(1) = "50100" And AWE_65.Value(2) = "20240")
    

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next

    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Uncategorized")
Private Sub TestMethod_ReadFromQuery_OperatorList_TwentyDepts()
    On Error GoTo TestFail
    
    'Arrange:
    Dim AWE As ApprovalWorkflowConfig
    Dim DeptCollection As DepartmentCollection
    Dim Departments(19) As Department
    Dim Index As Long
    
    Set AWE = New ApprovalWorkflowConfig
    Set DeptCollection = New DepartmentCollection
    
    Index = 0
    Do While Index <= 19
        Set Departments(Index) = New Department
        Index = Index + 1
    Loop
    
    Departments(0).DeptID = "20240"
    Departments(1).DeptID = "50001"
    Departments(2).DeptID = "50100"
    Departments(3).DeptID = "50102"
    Departments(4).DeptID = "50110"
    Departments(5).DeptID = "50120"
    Departments(6).DeptID = "50130"
    Departments(7).DeptID = "50140"
    Departments(8).DeptID = "50200"
    Departments(9).DeptID = "50210"
    Departments(10).DeptID = "50220"
    Departments(11).DeptID = "50230"
    Departments(12).DeptID = "50340"
    Departments(13).DeptID = "50341"
    Departments(14).DeptID = "50501"
    Departments(15).DeptID = "50510"
    Departments(16).DeptID = "50520"
    Departments(17).DeptID = "50530"
    Departments(18).DeptID = "50540"
    Departments(19).DeptID = "50550"
    
    Index = 0
    Do While Index <= 19
        DeptCollection.Add Departments(Index)
        Index = Index + 1
    Loop
    
    'Act:
    AWE.ReadFrom_QFS_SEC_EOAW_APPROVAL_SETUP_sheet wbAWE.Sheets.Item(1), 65, DeptCollection
    
    'Assert:
    Index = 0
    Do While Index <= 19
        Assert.IsTrue Departments(Index).DeptID = AWE.Value(Index + 1)
        Index = Index + 1
    Loop
    Assert.IsTrue AWE.ValueCount = 20

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Uncategorized")
Private Sub TestMethod_ReadFromQuery_OperatorList_21Depts_OneExcluded()
    On Error GoTo TestFail
    
    'Arrange:
    Dim AWE As ApprovalWorkflowConfig
    Dim DeptCollection As DepartmentCollection
    Dim Departments(20) As Department
    Dim Index As Long
    
    Set AWE = New ApprovalWorkflowConfig
    Set DeptCollection = New DepartmentCollection
    
    Index = 0
    Do While Index <= 20
        Set Departments(Index) = New Department
        Index = Index + 1
    Loop
    
    Departments(0).DeptID = "20240"
    Departments(1).DeptID = "50001"
    Departments(2).DeptID = "50100"
    Departments(3).DeptID = "50102"
    Departments(4).DeptID = "50110"
    Departments(5).DeptID = "50120"
    Departments(6).DeptID = "50130"
    Departments(7).DeptID = "50140"
    Departments(8).DeptID = "50200"
    Departments(9).DeptID = "50210"
    Departments(10).DeptID = "50220"
    Departments(11).DeptID = "50230"
    Departments(12).DeptID = "50340"
    Departments(13).DeptID = "50341"
    Departments(14).DeptID = "50501"
    Departments(15).DeptID = "50510"
    Departments(16).DeptID = "50520"
    Departments(17).DeptID = "50530"
    Departments(18).DeptID = "50540"
    Departments(19).DeptID = "50550"
    Departments(20).DeptID = "FALSE"
    
    Index = 0
    Do While Index <= 20
        DeptCollection.Add Departments(Index)
        Index = Index + 1
    Loop
    
    'Act:
    AWE.ReadFrom_QFS_SEC_EOAW_APPROVAL_SETUP_sheet wbAWE.Sheets.Item(1), 65, DeptCollection
    
    'Assert:
    Index = 0
    Do While Index < AWE.ValueCount
        Assert.IsTrue Departments(Index).DeptID = AWE.Value(Index + 1)
        Index = Index + 1
    Loop
    Assert.IsTrue AWE.ValueCount = 20

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("No Fail")
Private Sub TestMethod_ReadFromWorksheet_NoFail()
    On Error GoTo TestFail
    
    'Arrange:
    Dim AWEColl As AWConfigCollection
    Dim DisplayAlerts As Boolean
    Dim wbDepartments As Workbook
    Dim Departments As DepartmentCollection
    
    Set AWEColl = New AWConfigCollection
    Set Departments = New DepartmentCollection
    Set wbDepartments = Workbooks _
        .Open(ThisWorkbook.Path & "/test_data/ALL_DEPTS_BY_SETID_ANON.csv", ReadOnly:=True)
    
    Departments.AddDepartmentsFromWorksheet wbDepartments.Sheets.Item(1)
    DisplayAlerts = Application.DisplayAlerts
    Application.DisplayAlerts = False
    wbDepartments.Close SaveChanges:=False
    Application.DisplayAlerts = DisplayAlerts
    
    'Act:
    ' Query header begins on row 2
    AWEColl.ReadFromWorksheet wbAWE.Sheets.Item(1), Departments, 2
    
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
Private Sub TestMethod_ReadFromWorksheet_CountCorrect()
    On Error GoTo TestFail
    
    'Arrange:
    Dim AWEColl As AWConfigCollection
    Dim DisplayAlerts As Boolean
    Dim wbDepartments As Workbook
    Dim Departments As DepartmentCollection
    
    Set AWEColl = New AWConfigCollection
    Set Departments = New DepartmentCollection
    Set wbDepartments = Workbooks _
        .Open(ThisWorkbook.Path & "/test_data/ALL_DEPTS_BY_SETID_ANON.csv", ReadOnly:=True)
    
    Departments.AddDepartmentsFromWorksheet wbDepartments.Sheets.Item(1)
    DisplayAlerts = Application.DisplayAlerts
    Application.DisplayAlerts = False
    wbDepartments.Close SaveChanges:=False
    Application.DisplayAlerts = DisplayAlerts
    
    'Act:
    ' Query header begins on row 2
    AWEColl.ReadFromWorksheet wbAWE.Sheets.Item(1), Departments, 2
    
    'Assert:
    Assert.IsTrue AWEColl.Count = 158

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Uncategorized")
Private Sub TestMethod_ReadFromWorksheet_FirstRowCorrect()
    On Error GoTo TestFail
    
    'Arrange:
    Dim AWEColl As AWConfigCollection
    Dim DisplayAlerts As Boolean
    Dim wbDepartments As Workbook
    Dim Departments As DepartmentCollection
    
    Set AWEColl = New AWConfigCollection
    Set Departments = New DepartmentCollection
    Set wbDepartments = Workbooks _
        .Open(ThisWorkbook.Path & "/test_data/ALL_DEPTS_BY_SETID_ANON.csv", ReadOnly:=True)
    
    Departments.AddDepartmentsFromWorksheet wbDepartments.Sheets.Item(1)
    DisplayAlerts = Application.DisplayAlerts
    Application.DisplayAlerts = False
    wbDepartments.Close SaveChanges:=False
    Application.DisplayAlerts = DisplayAlerts
    
    'Act:
    ' Query header begins on row 2
    AWEColl.ReadFromWorksheet wbAWE.Sheets.Item(1), Departments, 2
    
    'Assert:
    With AWEColl.Item(1)
        Assert.IsTrue .ProcessID = "PurchaseOrder"
        Assert.IsTrue .DefinitionID = "SHARE"
        Assert.IsTrue .Description = "Line Fiscal"
        Assert.IsTrue .EffectiveDate = #1/1/1902#
        Assert.IsTrue .EffectiveStatus = "I"
        Assert.IsTrue .Stage = 10
        Assert.IsTrue .Path = 1
        Assert.IsTrue .Step = 1
        Assert.IsTrue .Approvers = "Supervisor by UserId"
        Assert.IsTrue .StepCriteriaDescription = "Step Criteria Definition for Line Level"
        Assert.IsTrue .StepFieldName = vbNullString
        Assert.IsTrue .ValueCount = 0
    End With
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Uncategorized")
Private Sub TestMethod_ReadFromWorksheet_LastRowCorrect()
    On Error GoTo TestFail
    
    'Arrange:
    Dim AWEColl As AWConfigCollection
    Dim DisplayAlerts As Boolean
    Dim wbDepartments As Workbook
    Dim Departments As DepartmentCollection
    
    Set AWEColl = New AWConfigCollection
    Set Departments = New DepartmentCollection
    Set wbDepartments = Workbooks _
        .Open(ThisWorkbook.Path & "/test_data/ALL_DEPTS_BY_SETID_ANON.csv", ReadOnly:=True)
    
    Departments.AddDepartmentsFromWorksheet wbDepartments.Sheets.Item(1)
    DisplayAlerts = Application.DisplayAlerts
    Application.DisplayAlerts = False
    wbDepartments.Close SaveChanges:=False
    Application.DisplayAlerts = DisplayAlerts
    
    'Act:
    ' Query header begins on row 2
    AWEColl.ReadFromWorksheet wbAWE.Sheets.Item(1), Departments, 2
    
    'Assert:
    With AWEColl.Item(AWEColl.Count)
        Assert.IsTrue .ProcessID = "VoucherApproval"
        Assert.IsTrue .DefinitionID = "WA999"
        Assert.IsTrue .Description = "WA999 Voucher AWE"
        Assert.IsTrue .EffectiveDate = #1/4/1901#
        Assert.IsTrue .EffectiveStatus = "A"
        Assert.IsTrue .Stage = 30
        Assert.IsTrue .Path = 1
        Assert.IsTrue .Step = 1
        Assert.IsTrue .Approvers = "CTC_UL_VCHR_AP_REVIEW"
        Assert.IsTrue .StepCriteriaDescription = vbNullString
        Assert.IsTrue .StepFieldName = vbNullString
        Assert.IsTrue .ValueCount = 0
    End With
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


