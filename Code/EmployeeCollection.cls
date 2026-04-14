VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EmployeeCollection"
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

Option Explicit
'@Folder("PSSec")

Private Type TEmployeeCollection
    Employees As Collection ' Collection of Employees
End Type
Private this As TEmployeeCollection

Private Sub Class_Initialize()
    Set this.Employees = New Collection
End Sub

Private Sub Class_Terminate()
    Set this.Employees = Nothing
End Sub

Public Sub Add(ByVal Emp As Employee)
    Dim Index As Long
    Dim EmpIndex As Long
    Dim Flag As Boolean
    Flag = False ' false by default
    For Index = 1 To Count()
        With Item(Index)
            If .EmplID = Emp.EmplID Then
                Flag = True
                For EmpIndex = 1 To Emp.RoleCount()
                    .AddRole Emp.Role(EmpIndex)
                Next EmpIndex
            End If
        End With
    Next Index
    If Not Flag Then
        this.Employees.Add Emp
    End If
End Sub

Public Sub Remove(ByVal Index As Long)
    this.Employees.Remove Index
End Sub

Public Function Count() As Long
    Count = this.Employees.Count
End Function

Public Function Item(ByVal Index As Long) As Employee
    Set Item = this.Employees.Item(Index)
End Function

'@Description "Add Employee objects created from the output of PSQuery QFS_SEC_USER_ROLES_BY_UNIT"
Public Sub ReadFromWorksheet(ByVal QFS_SEC_USER_ROLES_BY_UNIT As Worksheet)
Attribute ReadFromWorksheet.VB_Description = "Add Employee objects created from the output of PSQuery QFS_SEC_USER_ROLES_BY_UNIT"
    Dim Emp As Employee
    Dim Index As Long
    ' First non-header row is 3
    With QFS_SEC_USER_ROLES_BY_UNIT
        For Index = 3 To .UsedRange.Rows.Count
            Set Emp = New Employee
            Emp.EmplID = .Range("C" & Index).Value2
            Emp.Name = .Range("D" & Index).Value2
            Emp.HRStatus = .Range("K" & Index).Value2
            Emp.AddRole .Range("G" & Index).Value2
            Add Emp
        Next Index
    End With
End Sub
