VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DepartmentCollection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
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

'@Folder("GL")
Option Explicit

Private Type TDepartmentCollection
    Departments As Collection
End Type
Private this As TDepartmentCollection

Private Sub Class_Initialize()
    Set this.Departments = New Collection
End Sub

Public Sub Add(ByVal Item As Department)
    this.Departments.Add Item
End Sub

Public Sub Remove(ByVal Index As Variant)
    this.Departments.Remove Index
End Sub

Public Function Count() As Long
    Count = this.Departments.Count
End Function

Public Function Item(ByVal Index As Variant) As Department
    Set Item = this.Departments.Item(Index)
End Function

Public Function Filter(Optional ByVal EmplID As String = vbNullString) As DepartmentCollection
    Dim FilteredDepartmentCollection As DepartmentCollection
    Dim Dept As Variant ' to be used in for each loop
    Set FilteredDepartmentCollection = New DepartmentCollection

    For Each Dept In this.Departments
        If (EmplID = vbNullString Or Dept.ManagerID = EmplID) Then
            FilteredDepartmentCollection.Add Dept
        End If
    Next Dept

    Set Filter = FilteredDepartmentCollection
End Function

' TODO: Rename
Public Sub ReadFromWorksheet(ByVal ws As Worksheet)
    Dim RowIndex As Long
    Dim Dept As Department

    RowIndex = 3
    Do While RowIndex <= ws.Rows.Count
        Set Dept = New Department
        Dept.ReadDepartmentFromWorksheet ws, RowIndex
        Add Dept
        Set Dept = Nothing
        RowIndex = RowIndex + 1
    Loop
End Sub
