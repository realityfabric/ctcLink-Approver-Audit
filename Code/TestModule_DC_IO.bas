Attribute VB_Name = "TestModule_DC_IO"
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
Private Fakes As Object
Private TestBook As Workbook
Private TestSheet As Worksheet
Private TestDataFolderPath As String

'@ModuleInitialize
Private Sub ModuleInitialize()
    'this method runs once per module.
    Set Assert = CreateObject("Rubberduck.AssertClass")
    Set Fakes = CreateObject("Rubberduck.FakesProvider")

    TestDataFolderPath = ThisWorkbook.Path & "/test_data"
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
    Set Assert = Nothing
    Set Fakes = Nothing
End Sub

'@TestInitialize
Private Sub TestInitialize()
    Set TestBook = Workbooks.Open(TestDataFolderPath & "/ALL_DEPTS_BY_SETID_ANON.csv", ReadOnly:=True)
    Set TestSheet = TestBook.Sheets.Item(1)
End Sub

'@TestCleanup
Private Sub TestCleanup()
    TestBook.Close SaveChanges:=False
    Set TestBook = Nothing
    Set TestSheet = Nothing
End Sub

'@TestMethod("No Fail")
Private Sub TestMethod_AddDepartmentCollectionFromSheet_NoFail()
    On Error GoTo TestFail
    
    'Arrange:
    Dim Departments As DepartmentCollection
    Set Departments = New DepartmentCollection
    
    'Act:
    Departments.AddDepartmentsFromWorksheet TestSheet
    
    'Assert:
    Assert.Succeed
    
    Set Departments = Nothing

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("No Fail")
Private Sub TestMethod_AddDepartmentCollectionFromSheet_Count()
    On Error GoTo TestFail
    
    'Arrange:
    Dim Departments As DepartmentCollection
    Set Departments = New DepartmentCollection
    
    'Act:
    Departments.AddDepartmentsFromWorksheet TestSheet
    
    'Assert:
    Assert.IsTrue Departments.Count = 773
    
    Set Departments = Nothing

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

