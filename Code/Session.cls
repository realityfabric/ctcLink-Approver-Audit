VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Session"
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

'@Folder("ApproverValidation")
Option Explicit

Private Type TSession
   FormClosedWithoutRunning As Boolean
   fUserRoles As String ' QFS_SEC_USER_ROLES_BY_UNIT
   fDepartments As String ' ALL_DEPTS_BY_SETID
   fExpenseApprovers As String ' QFS_SEC_OPR_EXP_APPRVR
   fApprovalSetup As String ' QFS_SEC_EOAW_APPROVAL_SETUP
End Type
Private this As TSession

Private Sub Class_Initialize()
    this.FormClosedWithoutRunning = True
End Sub

Public Property Let FormClosedWithoutRunning(ByVal Value As Boolean)
    this.FormClosedWithoutRunning = Value
End Property

Public Property Get FormClosedWithoutRunning() As Boolean
    FormClosedWithoutRunning = this.FormClosedWithoutRunning
End Property

Public Property Let fUserRoles(ByVal Value As String)
    this.fUserRoles = Value
End Property

Public Property Get fUserRoles() As String
    fUserRoles = this.fUserRoles
End Property

Public Property Let fDepartments(ByVal Value As String)
    this.fDepartments = Value
End Property

Public Property Get fDepartments() As String
    fDepartments = this.fDepartments
End Property

Public Property Let fExpenseApprovers(ByVal Value As String)
    this.fExpenseApprovers = Value
End Property

Public Property Get fExpenseApprovers() As String
    fExpenseApprovers = this.fExpenseApprovers
End Property

Public Property Let fApprovalSetup(ByVal Value As String)
    this.fApprovalSetup = Value
End Property

Public Property Get fApprovalSetup() As String
    fApprovalSetup = this.fApprovalSetup
End Property
