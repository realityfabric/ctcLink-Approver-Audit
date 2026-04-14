Attribute VB_Name = "Tests_Employee_Properties"
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
'@Folder("Tests.Properties")

Option Explicit
Option Private Module

Private Assert As Object
Private Emp As Employee

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
    Set Emp = New Employee
End Sub

'@TestCleanup
Private Sub TestCleanup()
    Set Emp = Nothing
End Sub

'@TestMethod("Letter")
Private Sub TestMethod_LetEmplID_NoFail()
    On Error GoTo TestFail
    
    'Arrange:
    'Act:
    Emp.EmplID = "100000000"
    
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

'@TestMethod("Getter")
Private Sub TestMethod_GetEmplID_Correct()
    On Error GoTo TestFail
    
    'Arrange:
    Emp.EmplID = "100000000"
    
    'Act:
    'Assert:
    Assert.IsTrue "100000000" = Emp.EmplID

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Letter")
Private Sub TestMethod_LetName_NoFail()
    On Error GoTo TestFail
    
    'Arrange:
    'Act:
    Emp.Name = "Doe, Jane"
    
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

'@TestMethod("Getter")
Private Sub TestMethod_GetName_Correct()
    On Error GoTo TestFail
    
    'Arrange:
    Emp.Name = "Doe, Jane"
    
    'Act:
    'Assert:
    Assert.IsTrue "Doe, Jane" = Emp.Name

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Letter")
Private Sub TestMethod_LetHRStatusA_NoFail()
    On Error GoTo TestFail
    
    'Arrange:
    'Act:
    Emp.HRStatus = "A"
    
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

'@TestMethod("Letter")
Private Sub TestMethod_LetHRStatusI_NoFail()
    On Error GoTo TestFail
    
    'Arrange:
    'Act:
    Emp.HRStatus = "I"
    
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

'@TestMethod("Getter")
Private Sub TestMethod_GetHRStatusA_Correct()
    On Error GoTo TestFail
    
    'Arrange:
    Emp.HRStatus = "A"
    
    'Act:
    'Assert:
    Assert.IsTrue "A" = Emp.HRStatus

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Getter")
Private Sub TestMethod_GetHRStatusI_Correct()
    On Error GoTo TestFail
    
    'Arrange:
    Emp.HRStatus = "I"
    
    'Act:
    'Assert:
    Assert.IsTrue "I" = Emp.HRStatus

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Uncategorized")
Private Sub TestMethod_IsActive_True()
    On Error GoTo TestFail
    
    'Arrange:
    'Act:
    Emp.HRStatus = "A"
    
    'Assert:
    Assert.IsTrue Emp.IsActive()

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Uncategorized")
Private Sub TestMethod_IsActive_False()
    On Error GoTo TestFail
    
    'Arrange:
    'Act:
    Emp.HRStatus = "I"
    
    'Assert:
    Assert.IsFalse Emp.IsActive()

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

