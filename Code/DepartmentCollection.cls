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

'@Description "Add Department instances to this object from a worksheet containing the ALL_DEPTS_BY_SETID query."
Public Sub AddDepartmentsFromWorksheet(ByVal ws_ALL_DEPTS_BY_SETID As Worksheet)
Attribute AddDepartmentsFromWorksheet.VB_Description = "Add Department instances to this object from a worksheet containing the ALL_DEPTS_BY_SETID query."
    Dim RowIndex As Long
    Dim Dept As Department

    RowIndex = 3
    Do While RowIndex <= ws_ALL_DEPTS_BY_SETID.UsedRange.Rows.Count
        Set Dept = New Department
        Dept.ReadFromWorksheet ws_ALL_DEPTS_BY_SETID, RowIndex
        Add Dept
        Set Dept = Nothing
        RowIndex = RowIndex + 1
    Loop
End Sub

'@Description "Return a DepartmentCollection with Departments not matching EXAPPROVERS
Public Function DepartmentsWithExpenseApproverMismatch(ByVal ExpenseApprovals As ExpenseApprovalCollection) As DepartmentCollection
    Dim Departments As DepartmentCollection
    Dim FilteredExApprovers As ExpenseApprovalCollection
    Dim Dept As Variant ' Department, used in For Each loop
    Dim Index As Long
    Set Departments = New DepartmentCollection
    Set FilteredExApprovers = ExpenseApprovals.Filter(ApproverType:="EXAPPROVER")
    
    For Each Dept In this.Departments
        Index = 1
        Do While Index <= FilteredExApprovers.Count
            If FilteredExApprovers.Item(Index).DepartmentManagerMismatch(Dept) Then
                Departments.Add Dept
            End If
            Index = Index + 1
        Loop
    Next Dept
    
    Set DepartmentsWithExpenseApproverMismatch = Departments
End Function

'@Description "Return a DepartmentCollection with Departments not covered in an ExpenseApprovalCollection"
Public Function DepartmentsWithoutExpenseApproval(ByVal ExpenseApprovals As ExpenseApprovalCollection) As DepartmentCollection
Attribute DepartmentsWithoutExpenseApproval.VB_Description = "Return a DepartmentCollection with Departments not covered in an ExpenseApprovalCollection"
    Dim FilteredExApprovers As ExpenseApprovalCollection
    Dim Dept As Variant ' Department, used in For Each loop
    Dim Index As Long
    Dim IsDeptCovered As Boolean
    Set DepartmentsWithoutExpenseApproval = New DepartmentCollection
    Set FilteredExApprovers = ExpenseApprovals.Filter(ApproverType:="EXAPPROVER")

    For Each Dept In this.Departments
        Index = 1
        IsDeptCovered = False ' Default to false
        Do While Index <= FilteredExApprovers.Count
            ' If there is at least one ExpenseApproval that covers the department then Dept is covered
            If FilteredExApprovers.Item(Index).DepartmentInRange(Dept.DeptID) Then IsDeptCovered = True
            
            Index = Index + 1
        Loop
        If Not IsDeptCovered Then DepartmentsWithoutExpenseApproval.Add Dept
    Next Dept
    
    ' Return DepartmentsWithoutExpenseApproval
End Function
