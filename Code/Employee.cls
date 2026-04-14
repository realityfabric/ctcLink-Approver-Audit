VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Employee"
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

Private Type TEmployee
    EmplID As String
    Name As String
    SecurityRoles As Collection
    HRStatus As String
    Departments As DepartmentCollection
    ExpenseApprovals As ExpenseApprovalCollection
End Type
Private this As TEmployee

Private Sub Class_Initialize()
    Set this.SecurityRoles = New Collection
    Set this.Departments = New DepartmentCollection
    Set this.ExpenseApprovals = New ExpenseApprovalCollection
End Sub

Public Property Get EmplID() As String
    EmplID = this.EmplID
End Property

Public Property Let EmplID(ByVal Value As String)
    this.EmplID = Value
End Property

Public Property Get Name() As String
    Name = this.Name
End Property

Public Property Let Name(ByVal Value As String)
    this.Name = Value
End Property

Public Property Get HRStatus() As String
    HRStatus = this.HRStatus
End Property

Public Property Let HRStatus(ByVal Value As String)
    this.HRStatus = Value
End Property

'@Description "Check if HRStatus is 'A'"
Public Function IsActive() As Boolean
Attribute IsActive.VB_Description = "Check if HRStatus is 'A'"
    IsActive = (HRStatus = "A")
End Function

'@Description "If the Employee does not already have the role, add the role to SecurityRoles
Public Sub AddRole(ByVal Role As String)
    If Not HasRole(Role) Then
        this.SecurityRoles.Add Role
    End If
End Sub

Public Sub RemoveRole(ByVal Index As Long)
    this.SecurityRoles.Remove Index
End Sub

Public Function RoleCount() As Long
    RoleCount = this.SecurityRoles.Count
End Function

'@Description "Access the employee's Security Roles"
Public Function Role(ByVal Index As Long) As String
Attribute Role.VB_Description = "Access the employee's Security Roles"
    Role = this.SecurityRoles.Item(Index)
End Function

'@Description "Return the Department at Index"
Public Function Department(ByVal Index As Long) As Department
Attribute Department.VB_Description = "Return the Department at Index"
    Set Department = this.Departments.Item(Index)
End Function

'@Description "Return the ExpenseApproval at Index"
Public Function ExpenseApproval(ByVal Index As Long) As ExpenseApproval
Attribute ExpenseApproval.VB_Description = "Return the ExpenseApproval at Index"
    Set ExpenseApproval = this.ExpenseApprovals.Item(Index)
End Function

'@Description "Add Department if its ManagerID matches the employee's EmplID"
Public Sub AddDepartment(ByVal Dept As Department)
Attribute AddDepartment.VB_Description = "Add Department if its ManagerID matches the employee's EmplID"
    If Dept.ManagerID = EmplID() Then this.Departments.Add Dept
End Sub

'@Description "Return the count of Departments"
Public Function DepartmentCount() As Long
Attribute DepartmentCount.VB_Description = "Return the count of Departments"
    DepartmentCount = this.Departments.Count
End Function

'@Description "Add ExpenseApproval if its EmplID matches the employee's EmplID"
Public Sub AddExpenseApproval(ByVal ExpAppr As ExpenseApproval)
Attribute AddExpenseApproval.VB_Description = "Add ExpenseApproval if its EmplID matches the employee's EmplID"
    If ExpAppr.EmplID = EmplID() Then this.ExpenseApprovals.Add ExpAppr
End Sub

'@Description "Return the count of ExpenseApprovals"
Public Function ExpenseApprovalCount() As Long
Attribute ExpenseApprovalCount.VB_Description = "Return the count of ExpenseApprovals"
    ExpenseApprovalCount = this.ExpenseApprovals.Count
End Function

Public Function HasRole(ByVal Role As String) As Boolean
    Dim Index As Long
    HasRole = False
    For Index = 1 To this.SecurityRoles.Count
        If this.SecurityRoles.Item(Index) = Role Then
            HasRole = True
        End If
    Next Index
End Function

'@Description "Replace this Employee's DepartmentCollection with a different DepartmentCollection."
'@Ignore WriteOnlyProperty
Public Property Set Departments(ByVal DCollection As DepartmentCollection)
Attribute Departments.VB_Description = "Replace this Employee's DepartmentCollection with a different DepartmentCollection."
    Set this.Departments = DCollection
End Property
