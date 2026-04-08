Attribute VB_Name = "TestModule_AWConfig_Properties"
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
Private Config As ApprovalWorkflowConfig

'@ModuleInitialize
Private Sub ModuleInitialize()
    'this method runs once per module.
    Set Assert = CreateObject("Rubberduck.AssertClass")
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
    'this method runs once per module.
    Set Assert = Nothing
End Sub

'@TestInitialize
Private Sub TestInitialize()
    Set Config = New ApprovalWorkflowConfig
End Sub

'@TestCleanup
Private Sub TestCleanup()
    Set Config = Nothing
End Sub

'@TestMethod("Letter")
Private Sub Let_Description_NoFail()
    On Error GoTo TestFail
    
    'Arrange:
    'Act:
    Config.Description = "ApprovalWorkflowConfig Description Test"
    
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
Private Sub Get_Description_NoFail()
    On Error GoTo TestFail
    
    'Arrange:
    '@Ignore VariableNotUsed
    Dim ConfigDescr As String
    Config.Description = "ApprovalWorkflowConfig Description Test"

    'Act:
    '@Ignore AssignmentNotUsed
    ConfigDescr = Config.Description

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
Private Sub Get_Description_DescriptionCorrect()
    On Error GoTo TestFail
    
    'Arrange:
    Config.Description = "ApprovalWorkflowConfig Description Test"

    'Act:
    'Assert:
    Assert.IsTrue "ApprovalWorkflowConfig Description Test" = Config.Description

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Letter")
Private Sub Let_Stage_NoFail()
    On Error GoTo TestFail
    
    'Arrange:
    'Act:
    Config.Stage = 10
    
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
Private Sub Get_Stage_NoFail()
    On Error GoTo TestFail
    
    'Arrange:
    '@Ignore VariableNotUsed
    Dim ConfigStage As Long
    Config.Stage = 10

    'Act:
    '@Ignore AssignmentNotUsed
    ConfigStage = Config.Stage

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
Private Sub Get_Stage_StageCorrect()
    On Error GoTo TestFail
    
    'Arrange:
    Config.Stage = 10

    'Act:
    'Assert:
    Assert.IsTrue 10 = Config.Stage

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Letter")
Private Sub Let_Path_NoFail()
    On Error GoTo TestFail
    
    'Arrange:
    'Act:
    Config.Path = 10
    
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
Private Sub Get_Path_NoFail()
    On Error GoTo TestFail
    
    'Arrange:
    '@Ignore VariableNotUsed
    Dim ConfigPath As Long
    Config.Path = 10

    'Act:
    '@Ignore AssignmentNotUsed
    ConfigPath = Config.Path

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
Private Sub Get_Path_PathCorrect()
    On Error GoTo TestFail
    
    'Arrange:
    Config.Path = 10

    'Act:
    'Assert:
    Assert.IsTrue 10 = Config.Path

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Letter")
Private Sub Let_Step_NoFail()
    On Error GoTo TestFail
    
    'Arrange:
    'Act:
    Config.Step = 10
    
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
Private Sub Get_Step_NoFail()
    On Error GoTo TestFail
    
    'Arrange:
    '@Ignore VariableNotUsed
    Dim ConfigStep As Long
    Config.Step = 10

    'Act:
    '@Ignore AssignmentNotUsed
    ConfigStep = Config.Step

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
Private Sub Get_Step_StepCorrect()
    On Error GoTo TestFail
    
    'Arrange:
    Config.Step = 10

    'Act:
    'Assert:
    Assert.IsTrue 10 = Config.Step

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Letter")
Private Sub Let_Approvers_NoFail()
    On Error GoTo TestFail
    
    'Arrange:
    'Act:
    Config.Approvers = "APPROVERS TEST"
    
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
Private Sub Get_Approvers_NoFail()
    On Error GoTo TestFail
    
    'Arrange:
    '@Ignore VariableNotUsed
    Dim ConfigApprovers As String
    Config.Approvers = "APPROVERS TEST"

    'Act:
    '@Ignore AssignmentNotUsed
    ConfigApprovers = Config.Approvers

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
Private Sub Get_Approvers_ApproversCorrect()
    On Error GoTo TestFail
    
    'Arrange:
    Config.Approvers = "APPROVERS TEST"

    'Act:
    'Assert:
    Assert.IsTrue "APPROVERS TEST" = Config.Approvers

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Letter")
Private Sub Let_StepCriteriaDescription_NoFail()
    On Error GoTo TestFail
    
    'Arrange:
    'Act:
    Config.StepCriteriaDescription = "StepCriteriaDescription Test"
    
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
Private Sub Get_StepCriteriaDescription_NoFail()
    On Error GoTo TestFail
    
    'Arrange:
    '@Ignore VariableNotUsed
    Dim ConfigStepCriteriaDescription As String
    Config.StepCriteriaDescription = "StepCriteriaDescription Test"

    'Act:
    '@Ignore AssignmentNotUsed
    ConfigStepCriteriaDescription = Config.StepCriteriaDescription

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
Private Sub Get_StepCriteriaDescription_StepCriteriaDescriptionCorrect()
    On Error GoTo TestFail
    
    'Arrange:
    Config.StepCriteriaDescription = "StepCriteriaDescription Test"

    'Act:
    'Assert:
    Assert.IsTrue "StepCriteriaDescription Test" = Config.StepCriteriaDescription

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Letter")
Private Sub Let_EffectiveDate_NoFail()
    On Error GoTo TestFail
    
    'Arrange:
    'Act:
    Config.EffectiveDate = #1/1/2026#
    
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
Private Sub Get_EffectiveDate_NoFail()
    On Error GoTo TestFail
    
    'Arrange:
    '@Ignore VariableNotUsed
    Dim ConfigEffectiveDate As Date
    Config.EffectiveDate = #1/1/2026#

    'Act:
    '@Ignore AssignmentNotUsed
    ConfigEffectiveDate = Config.EffectiveDate

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
Private Sub Get_EffectiveDate_EffectiveDateCorrect()
    On Error GoTo TestFail
    
    'Arrange:
    Config.EffectiveDate = #1/1/2026#

    'Act:
    'Assert:
    Assert.IsTrue #1/1/2026# = Config.EffectiveDate

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Letter")
Private Sub Let_EffectiveStatus_NoFail()
    On Error GoTo TestFail
    
    'Arrange:
    'Act:
    Config.EffectiveStatus = "A"
    
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
Private Sub Get_EffectiveStatus_NoFail()
    On Error GoTo TestFail
    
    'Arrange:
    '@Ignore VariableNotUsed
    Dim ConfigEffectiveStatus As String
    Config.EffectiveStatus = "A"

    'Act:
    '@Ignore AssignmentNotUsed
    ConfigEffectiveStatus = Config.EffectiveStatus

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
Private Sub Get_EffectiveStatus_EffectiveStatusCorrect()
    On Error GoTo TestFail
    
    'Arrange:
    Config.EffectiveStatus = "A"

    'Act:
    'Assert:
    Assert.IsTrue "A" = Config.EffectiveStatus

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub
