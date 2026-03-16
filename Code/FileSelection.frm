VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FileSelection 
   Caption         =   "FileSelection"
   ClientHeight    =   5055
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   9990.001
   OleObjectBlob   =   "FileSelection.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FileSelection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder("Forms")
Option Explicit

Private Sub BtnSelectApprovalSetup_Click()
    '@Ignore UseMeaningfulName
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    With fd
        .AllowMultiSelect = False
        .Title = "Select QFS_SEC_EOAW_APPROVAL_SETUP"
        If .Show = -1 Then
            Me.TextBoxApprovalSetup.Value = .SelectedItems.Item(1)
        End If
    End With
    Set fd = Nothing
End Sub

Private Sub BtnSelectDepartments_Click()
    '@Ignore UseMeaningfulName
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    With fd
        .AllowMultiSelect = False
        .Title = "Select ALL_DEPTS_BY_SETID"
        If .Show = -1 Then
            Me.TextBoxDepartments.Value = .SelectedItems.Item(1)
        End If
    End With
    Set fd = Nothing
End Sub

Private Sub BtnSelectExpenseApprovers_Click()
    '@Ignore UseMeaningfulName
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    With fd
        .AllowMultiSelect = False
        .Title = "Select QFS_SEC_OPR_EXP_APPRVR"
        If .Show = -1 Then
            Me.TextBoxExpenseApprovers.Value = .SelectedItems.Item(1)
        End If
    End With
    Set fd = Nothing
End Sub

Private Sub BtnSelectSecurityRoles_Click()
    '@Ignore UseMeaningfulName
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    With fd
        .AllowMultiSelect = False
        .Title = "Select QFS_SEC_USER_ROLES_BY_UNIT"
        If .Show = -1 Then
            Me.TextBoxSecurityRoles.Value = .SelectedItems.Item(1)
        End If
    End With
    Set fd = Nothing
End Sub

Private Sub BtnStart_Click()
    If Len(Me.TextBoxApprovalSetup.Value) > 0 And _
        Len(Me.TextBoxDepartments.Value) > 0 And _
        Len(Me.TextBoxExpenseApprovers.Value) > 0 And _
        Len(Me.TextBoxSecurityRoles.Value) > 0 _
        Then
        Main.Sesh.fApprovalSetup = Me.TextBoxApprovalSetup.Value
        Main.Sesh.fDepartments = Me.TextBoxDepartments.Value
        Main.Sesh.fExpenseApprovers = Me.TextBoxExpenseApprovers.Value
        Main.Sesh.fUserRoles = Me.TextBoxSecurityRoles.Value
        Main.Sesh.FormClosedWithoutRunning = False
        Me.Hide
    Else
        MsgBox "Please ensure that all files have been selected." _
        , vbExclamation
    End If
End Sub
