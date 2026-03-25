VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Department"
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

Private Type TDepartment
    DeptID As String
    Description As String
    ManagerID As String ' EmplID
End Type
Private this As TDepartment

'@Ignore ProcedureNotUsed
Public Property Let DeptID(ByVal Value As String)
    this.DeptID = Value
End Property

'@Ignore ProcedureNotUsed
Public Property Get DeptID() As String
    DeptID = this.DeptID
End Property

'@Ignore ProcedureNotUsed
Public Property Let Description(ByVal Value As String)
    this.Description = Value
End Property

'@Ignore ProcedureNotUsed
Public Property Get Description() As String
    Description = this.Description
End Property

'@Ignore ProcedureNotUsed
Public Property Let ManagerID(ByVal Value As String)
    this.ManagerID = Value
End Property

'@Ignore ProcedureNotUsed
Public Property Get ManagerID() As String
    ManagerID = this.ManagerID
End Property

'@Description "Set this object's fields based on a row from a worksheet containing the ALL_DEPTS_BY_SETID query."
'@Ignore ProcedureNotUsed
Public Sub ReadFromWorksheet(ByVal ws_ALL_DEPTS_BY_SETID As Worksheet, ByVal Row As Long)
Attribute ReadFromWorksheet.VB_Description = "Set this object's fields based on a row from a worksheet containing the ALL_DEPTS_BY_SETID query."
    Dim DataArray As Variant
    DataArray = ws_ALL_DEPTS_BY_SETID.Range("A" & Row & ":H" & Row).Value2
    this.DeptID = DataArray(1, 2)
    this.Description = DataArray(1, 5)
    this.ManagerID = DataArray(1, 8)
End Sub

'@Description "Check if a Department has at least one EXAPPROVER.
'@Ignore ProcedureNotUsed
Public Function DepartmentHasEXApproverInCollection(ByVal ExpenseApprovers As ExpenseApprovalCollection) As Boolean
    Dim flag As Boolean
    Dim Index As Long
    Index = 1

    ' Default to false
    flag = False
    Do While Index <= ExpenseApprovers.Count
        With ExpenseApprovers.Item(Index)
            If "EXAPPROVER" = .ApproverType And _
                .DepartmentInRange(DeptID) Then
                flag = True
            End If
        End With
        Index = Index + 1
    Loop

    DepartmentHasEXApproverInCollection = flag
End Function

'@Description "Check if a Department has at least one VPAPPROVER"
'@Ignore ProcedureNotUsed
Public Function DepartmentHasVPApproverInCollection(ByVal ExpenseApprovers As ExpenseApprovalCollection) As Boolean
    Dim flag As Boolean
    Dim Index As Long
    Index = 1

    'Default to false
    flag = False

    Do While Index <= ExpenseApprovers.Count
        With ExpenseApprovers.Item(Index)
            If "VPAPPROVER" = .ApproverType And _
                .DepartmentInRange(DeptID) Then
                flag = True
            End If
        End With
        Index = Index + 1
    Loop

    DepartmentHasVPApproverInCollection = flag
End Function

