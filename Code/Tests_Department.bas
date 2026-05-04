Attribute VB_Name = "Tests_Department"
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

'@ModuleInitialize
Private Sub ModuleInitialize()
    Set Assert = CreateObject("Rubberduck.AssertClass")
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
    Set Assert = Nothing
End Sub

'@TestMethod("Check")
Private Sub TestMethod_DepartmentHasManagerID()
    On Error GoTo TestFail
    
    'Arrange:
    Dim Dept As Department
    Set Dept = New Department
    Dept.DeptID = "1"
    Dept.ManagerID = "10"
    
    'Act:
    'Assert:
    Assert.IsTrue Dept.HasManager

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Check")
Private Sub TestMethod_DepartmentDoesNotHaveManagerID()
    On Error GoTo TestFail
    
    'Arrange:
    Dim Dept As Department
    Set Dept = New Department
    Dept.DeptID = "1"
    
    'Act:
    'Assert:
    Assert.IsFalse Dept.HasManager

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Check")
Private Sub ReadFromWorksheet_DepartmentWithLeadingZeroes()
    On Error GoTo TestFail
    
    'Arrange:
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets.Add()
    With ws
        .Range("A1").Value2 = "WA000" ' Business Unit
        .Range("B1").Value2 = "00001" ' Department ID
        .Range("E1").Value2 = "Leading 0s" ' Description
        .Range("H1").Value2 = "123456789" ' Manager ID
    End With
    
    Dim Dept As Department
    Set Dept = New Department
    Dept.ReadFromWorksheet ws, 1
    
    'Act:
    'Assert:
    'Assert.IsTrue Dept.DeptID = "00001"
    Assert.AreEqual "00001", Dept.DeptID

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    ' Record Value of DisplayAlerts, then disable DisplayAlerts
    ' Delete Worksheet, then set DisplayAlerts to previous value
    Dim DisplayAlerts As Boolean
    DisplayAlerts = Application.DisplayAlerts
    Application.DisplayAlerts = False
    ws.Delete
    Application.DisplayAlerts = DisplayAlerts
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub
