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
'@Folder("Tests")

Option Explicit
Option Private Module

Private Assert As Object
Private Fakes As Object

Private Dept1 As Department
Private Dept2 As Department
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
