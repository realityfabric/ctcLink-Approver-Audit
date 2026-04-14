Attribute VB_Name = "Tests_Employee_IO"
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
Private wbSecurityRoles As Workbook

'@ModuleInitialize
Private Sub ModuleInitialize()
    Set Assert = CreateObject("Rubberduck.AssertClass")
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
    Set Assert = Nothing
End Sub

'@TestInitialize
Private Sub TestInitialize()
    Set wbSecurityRoles = Workbooks.Open(ThisWorkbook.Path & "/test_data/QFS_SEC_USER_ROLES_BY_UNIT.csv", ReadOnly:=True)
End Sub

'@TestCleanup
Private Sub TestCleanup()
    wbSecurityRoles.Close SaveChanges:=False
    Set wbSecurityRoles = Nothing
End Sub

'@TestMethod("No Fail")
Private Sub EmployeeCollection_ReadFromWorksheet_NoFail()
    On Error GoTo TestFail
    
    'Arrange:
    Dim Employees As EmployeeCollection
    Set Employees = New EmployeeCollection
    
    'Act:
    Employees.ReadFromWorksheet wbSecurityRoles.Sheets.Item(1)
    
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
Private Sub EmployeeCollection_ReadFromWorksheet_CountCorrect()
    On Error GoTo TestFail
    
    'Arrange:
    Dim Employees As EmployeeCollection
    Set Employees = New EmployeeCollection
    
    'Act:
    Employees.ReadFromWorksheet wbSecurityRoles.Sheets.Item(1)
    
    'Assert:
    Assert.IsTrue Employees.Count = 4

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Uncategorized")
Private Sub Employee_ReadFromWorksheet_RoleCountCorrect()
    On Error GoTo TestFail
    
    'Arrange:
    Dim Employees As EmployeeCollection
    Set Employees = New EmployeeCollection
    
    'Act:
    Employees.ReadFromWorksheet wbSecurityRoles.Sheets.Item(1)
    
    'Assert:
    Dim Index As Long
    For Index = 1 To Employees.Count
        With Employees.Item(Index)
            If "100000000" = .EmplID Then
                Assert.IsTrue .RoleCount = 5
            ElseIf "200000000" = .EmplID Then
                Assert.IsTrue .RoleCount = 6
            ElseIf "300000000" = .EmplID Then
                Assert.IsTrue .RoleCount = 7
            ElseIf "400000000" = .EmplID Then
                Assert.IsTrue .RoleCount = 2
            Else
                Assert.Failure ' There are no other employees
            End If
        End With
    Next Index

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Uncategorized")
Private Sub Employee_ReadFromWorksheet_NamesCorrect()
    On Error GoTo TestFail
    
    'Arrange:
    Dim Employees As EmployeeCollection
    Set Employees = New EmployeeCollection
    
    'Act:
    Employees.ReadFromWorksheet wbSecurityRoles.Sheets.Item(1)
    
    'Assert:
    Dim Index As Long
    For Index = 1 To Employees.Count
        With Employees.Item(Index)
            If "100000000" = .EmplID Then
                Assert.IsTrue .Name = "Buffet, Jimmy"
            ElseIf "200000000" = .EmplID Then
                Assert.IsTrue .Name = "Jackson, Alan"
            ElseIf "300000000" = .EmplID Then
                Assert.IsTrue .Name = "Mercury, Fred"
            ElseIf "400000000" = .EmplID Then
                Assert.IsTrue .Name = "Neutron, Jimmy"
            Else
                Assert.Failure ' There are no other employees
            End If
        End With
    Next Index

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub
