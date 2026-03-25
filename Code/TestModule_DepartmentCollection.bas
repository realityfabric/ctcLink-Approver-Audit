Attribute VB_Name = "TestModule_DepartmentCollection"
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
'@Folder("Tests.GL")

Option Explicit
Option Private Module

Private Assert As Object
Private Fakes As Object

'@Ignore UseMeaningfulName
Private Dept1 As Department
'@Ignore UseMeaningfulName
Private Dept2 As Department
'@Ignore UseMeaningfulName
Private Dept3 As Department

'@ModuleInitialize
Private Sub ModuleInitialize()
    'this method runs once per module.
    Set Assert = CreateObject("Rubberduck.AssertClass")
    Set Fakes = CreateObject("Rubberduck.FakesProvider")
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
    'this method runs once per module.
    Set Assert = Nothing
    Set Fakes = Nothing
End Sub

'@TestInitialize
Private Sub TestInitialize()
    'This method runs before every test in the module..
    Set Dept1 = New Department
    Set Dept2 = New Department
    Set Dept3 = New Department
    
    Dept1.DeptID = "1"
    Dept2.DeptID = "2"
    Dept3.DeptID = "3"
    
    Dept1.Description = "Department 1"
    Dept2.Description = "Department 2"
    Dept3.Description = "Department 3"
    
    Dept1.ManagerID = "1"
    Dept2.ManagerID = "1"
    Dept3.ManagerID = "2"
End Sub

'@TestCleanup
Private Sub TestCleanup()
    'this method runs after every test in the module.
    
    Set Dept1 = Nothing
    Set Dept2 = Nothing
    Set Dept3 = Nothing
End Sub

'@TestMethod("No Fail")
Private Sub TestMethod_AddDepartments_NoFail()
    On Error GoTo TestFail
    
    'Arrange:
    '@Ignore UseMeaningfulName
    Dim DC As DepartmentCollection
    Set DC = New DepartmentCollection
    
    'Act:
    DC.Add Dept1
    DC.Add Dept2
    DC.Add Dept3
    
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

'@TestMethod("Wrapper")
'@Ignore UseMeaningfulName
Private Sub TestMethod_Count3()
    On Error GoTo TestFail
    
    'Arrange:
    '@Ignore UseMeaningfulName
    Dim DC As DepartmentCollection
    Set DC = New DepartmentCollection
    
    DC.Add Dept1
    DC.Add Dept2
    DC.Add Dept3
    
    'Act:
    'Assert:
    Assert.IsTrue DC.Count = 3

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("No Fail")
Private Sub TestMethod_Add3_Remove1_NoFail()
    On Error GoTo TestFail
    
    'Arrange:
    '@Ignore UseMeaningfulName
    Dim DC As DepartmentCollection
    Set DC = New DepartmentCollection
    
    'Act:
    DC.Add Dept1
    DC.Add Dept2
    DC.Add Dept3
    DC.Remove 1
    
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
'@Ignore UseMeaningfulName
Private Sub TestMethod_Add3_Remove1_Count2()
    On Error GoTo TestFail
    
    'Arrange:
    '@Ignore UseMeaningfulName
    Dim DC As DepartmentCollection
    Set DC = New DepartmentCollection
    
    'Act:
    DC.Add Dept1
    DC.Add Dept2
    DC.Add Dept3
    DC.Remove 1
    
    'Assert:
    Assert.IsTrue DC.Count = 2

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Filter")
'@Ignore UseMeaningfulName
Private Sub TestMethod_Add3_FilterToOne_Count1()
    On Error GoTo TestFail
    
    'Arrange:
    '@Ignore UseMeaningfulName
    Dim DC As DepartmentCollection
    Dim DCFiltered As DepartmentCollection
    Set DC = New DepartmentCollection
    
    DC.Add Dept1
    DC.Add Dept2
    DC.Add Dept3
    
    'Act:
    Set DCFiltered = DC.Filter(EmplID:="2")
    
    'Assert:
    Assert.IsTrue 1 = DCFiltered.Count()
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Filter")
'@Ignore UseMeaningfulName
Private Sub TestMethod_Add3_FilterToOne_ManagerID2()
    On Error GoTo TestFail
    
    'Arrange:
    '@Ignore UseMeaningfulName
    Dim DC As DepartmentCollection
    Dim DCFiltered As DepartmentCollection
    Set DC = New DepartmentCollection
    
    DC.Add Dept1
    DC.Add Dept2
    DC.Add Dept3
    
    'Act:
    Set DCFiltered = DC.Filter(EmplID:="2")
    
    'Assert:
    Assert.IsTrue DCFiltered.Item(1).ManagerID = "2"

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Filter")
'@Ignore UseMeaningfulName
Private Sub TestMethod_Add3_FilterTo2_ManagerID1()
    On Error GoTo TestFail
    
    'Arrange:
    '@Ignore UseMeaningfulName
    Dim DC As DepartmentCollection
    Dim DCFiltered As DepartmentCollection
    Set DC = New DepartmentCollection
    
    DC.Add Dept1
    DC.Add Dept2
    DC.Add Dept3
    
    'Act:
    Set DCFiltered = DC.Filter(EmplID:="1")
    
    'Assert:
    Assert.IsTrue DCFiltered.Count = 2
    Assert.IsTrue DCFiltered.Item(1).ManagerID = "1"
    Assert.IsTrue DCFiltered.Item(2).ManagerID = "1"
    Assert.IsTrue DCFiltered.Item(1).DeptID = "1"
    Assert.IsTrue DCFiltered.Item(2).DeptID = "2"

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Filter")
'@Ignore UseMeaningfulName
Private Sub TestMethod_Add3_FilterToOne_DeptID3()
    On Error GoTo TestFail
    
    'Arrange:
    '@Ignore UseMeaningfulName
    Dim DC As DepartmentCollection
    Dim DCFiltered As DepartmentCollection
    Set DC = New DepartmentCollection
    
    DC.Add Dept1
    DC.Add Dept2
    DC.Add Dept3
    
    'Act:
    Set DCFiltered = DC.Filter(EmplID:="2")
    
    'Assert:
    Assert.IsTrue "3" = DCFiltered.Item(1).DeptID

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


