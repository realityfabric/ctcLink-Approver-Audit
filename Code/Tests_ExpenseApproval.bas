Attribute VB_Name = "Tests_ExpenseApproval"
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
'@Folder("Tests")

Option Explicit
Option Private Module

Private Assert As Object

Private TestWorksheet As Worksheet

'@ModuleInitialize
Private Sub ModuleInitialize()
    Set Assert = CreateObject("Rubberduck.AssertClass")

    Set TestWorksheet = ThisWorkbook.Sheets.Add
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
    Dim DisplayAlerts As Boolean

    Set Assert = Nothing

    ' Save Application.DisplayAlerts value, set to False, delete the test sheet,
    '  then set Application.DisplayAlerts back to its previous value.
    DisplayAlerts = Application.DisplayAlerts
    Application.DisplayAlerts = False
    TestWorksheet.Delete
    Application.DisplayAlerts = DisplayAlerts
End Sub

'@TestInitialize
Private Sub TestInitialize()
    With TestWorksheet
        .Range("A1:H1").Value2 = Array( _
            "GL Unit", _
            "Approver Type", _
            "EmplID", _
            "Description", _
            "From Chartfield", _
            "To Chartfield", _
            "Last Name", _
            "First Name" _
        )

        .Range("A2:H2").Value2 = Array( _
            "WA190", _
            "EXAPPROVER", _
            "111111111", _
            "Test Department", _
            "10000", _
            "10500", _
            "John", _
            "Smith" _
        )
    End With
End Sub

'@TestCleanup
Private Sub TestCleanup()
    With TestWorksheet
        .UsedRange.Clear
    End With
End Sub

'@TestMethod("No Fail")
Private Sub TestMethod_ReadExpenseApprovalFromWorksheet_NoFail()
    On Error GoTo TestFail

    'Arrange:
    '@Ignore UseMeaningfulName
    Dim EA As ExpenseApproval
    Set EA = New ExpenseApproval

    'Act:
    EA.ReadExpenseApprovalFromWorksheet TestWorksheet, 2

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


'@TestMethod("Compare")
Private Sub TestMethod_DeptID_IsInRange_OneDeptID()
    On Error GoTo TestFail

    'Arrange:
    '@Ignore UseMeaningfulName
    Dim EA As ExpenseApproval
    Set EA = New ExpenseApproval

    EA.FromChartfield = "00000"
    EA.ToChartfield = "00000"

    'Act:
    'Assert:
    Assert.IsTrue EA.DepartmentInRange("00000")

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next

    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Compare")
Private Sub TestMethod_DeptID_IsNotInRange_1DeptID()
    On Error GoTo TestFail

    'Arrange:
    '@Ignore UseMeaningfulName
    Dim EA As ExpenseApproval
    Set EA = New ExpenseApproval

    EA.FromChartfield = "00000"
    EA.ToChartfield = "00000"

    'Act:
    'Assert:
    Assert.IsFalse EA.DepartmentInRange("00001")

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next

    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Compare")
Private Sub TestMethod_DeptID_IsInRange_2DeptID()
    On Error GoTo TestFail

    'Arrange:
    '@Ignore UseMeaningfulName
    Dim EA As ExpenseApproval
    Set EA = New ExpenseApproval

    EA.FromChartfield = "00000"
    EA.ToChartfield = "00001"

    'Act:
    'Assert:
    Assert.IsTrue EA.DepartmentInRange("00000")
    Assert.IsTrue EA.DepartmentInRange("00001")

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next

    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Compare")
Private Sub TestMethod_DeptID_IsNotInRange_2DeptID()
    On Error GoTo TestFail

    'Arrange:
    '@Ignore UseMeaningfulName
    Dim EA As ExpenseApproval
    Set EA = New ExpenseApproval

    EA.FromChartfield = "00000"
    EA.ToChartfield = "00001"

    'Act:
    'Assert:
    Assert.IsFalse EA.DepartmentInRange("00002")

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next

    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Compare")
Private Sub TestMethod_DeptID_IsInRange_AlphaNumeric()
    On Error GoTo TestFail

    'Arrange:
    '@Ignore UseMeaningfulName
    Dim EA As ExpenseApproval
    Set EA = New ExpenseApproval

    EA.FromChartfield = "ABC00"
    EA.ToChartfield = "ABC99"

    'Act:
    'Assert:
    Assert.IsTrue EA.DepartmentInRange("ABC00")
    Assert.IsTrue EA.DepartmentInRange("ABC99")

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next

    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Compare")
Private Sub TestMethod_DeptID_IsInRange_AlphaNumeric_NotFromOrTo()
    On Error GoTo TestFail

    'Arrange:
    '@Ignore UseMeaningfulName
    Dim EA As ExpenseApproval
    Set EA = New ExpenseApproval

    EA.FromChartfield = "ABC00"
    EA.ToChartfield = "ABC99"

    'Act:
    'Assert:
    Assert.IsTrue EA.DepartmentInRange("ABC25")
    Assert.IsTrue EA.DepartmentInRange("ABC50")
    Assert.IsTrue EA.DepartmentInRange("ABC75")

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next

    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


