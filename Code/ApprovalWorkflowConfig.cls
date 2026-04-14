VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ApprovalWorkflowConfig"
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

'@Folder("AWE")
Option Explicit

Private Type TApprovalWorkflowConfig
    ProcessID As String
    DefinitionID As String
    Description As String
    EffectiveDate As Date
    EffectiveStatus As String
    Stage As Long
    Path As Long
    Step As Long
    Approvers As String
    StepCriteriaDescription As String
    StepFieldName As String
    Values As Collection ' Will be a collection of strings
End Type
Private this As TApprovalWorkflowConfig

Private Sub Class_Initialize()
    Set this.Values = New Collection
End Sub

Private Sub Class_Terminate()
    Set this.Values = Nothing
End Sub

Public Property Get ProcessID() As String
    ProcessID = this.ProcessID
End Property

Public Property Let ProcessID(ByVal Value As String)
    this.ProcessID = Value
End Property

Public Property Get DefinitionID() As String
    DefinitionID = this.DefinitionID
End Property

Public Property Let DefinitionID(ByVal Value As String)
    this.DefinitionID = Value
End Property

Public Property Get Description() As String
    Description = this.Description
End Property

Public Property Let Description(ByVal Value As String)
    this.Description = Value
End Property

Public Property Get EffectiveDate() As Date
    EffectiveDate = this.EffectiveDate
End Property

Public Property Let EffectiveDate(ByVal Value As Date)
    this.EffectiveDate = Value
End Property

Public Property Get EffectiveStatus() As String
    EffectiveStatus = this.EffectiveStatus
End Property

Public Property Let EffectiveStatus(ByVal Value As String)
    this.EffectiveStatus = Value
End Property

Public Property Get Stage() As Long
    Stage = this.Stage
End Property

Public Property Let Stage(ByVal Value As Long)
    this.Stage = Value
End Property

Public Property Get Path() As Long
    Path = this.Path
End Property

Public Property Let Path(ByVal Value As Long)
    this.Path = Value
End Property

Public Property Get Step() As Long
    Step = this.Step
End Property

Public Property Let Step(ByVal Value As Long)
    this.Step = Value
End Property

Public Property Get Approvers() As String
    Approvers = this.Approvers
End Property

Public Property Let Approvers(ByVal Value As String)
    this.Approvers = Value
End Property

Public Property Get StepCriteriaDescription() As String
    StepCriteriaDescription = this.StepCriteriaDescription
End Property

Public Property Let StepCriteriaDescription(ByVal Value As String)
    this.StepCriteriaDescription = Value
End Property

Public Property Get StepFieldName() As String
    StepFieldName = this.StepFieldName
End Property

Public Property Let StepFieldName(ByVal Value As String)
    this.StepFieldName = Value
End Property

Public Property Get Values() As Collection
    Set Values = this.Values
End Property

Public Function Value(ByVal Index As Long) As String
    Value = this.Values.Item(Index)
End Function

Public Sub AddValue(ByVal Value As String)
    this.Values.Add Value
End Sub

Public Sub RemoveValue(ByVal Index As Long)
    this.Values.Remove (Index)
End Sub

Public Function ValueCount() As Long
    ValueCount = this.Values.Count
End Function

' Departments are required because I don't want to generate massive lists of department strings
' Which would be necessary if criteria was "Between" any department with a non-numeric value and another.
' By requiring Departments, we can simply iterate through the departments and add existing ones.
'@Ignore UseMeaningfulName, UnderscoreInPublicClassModuleMember
Public Sub ReadFrom_QFS_SEC_EOAW_APPROVAL_SETUP_sheet(ByVal ws As Worksheet, ByVal RowIndex As Long, ByVal Departments As DepartmentCollection)
    With ws
        ProcessID = .Range("A" & RowIndex).Value2
        DefinitionID = .Range("B" & RowIndex).Value2
        Description = .Range("C" & RowIndex).Value2
        EffectiveDate = .Range("D" & RowIndex).Value2
        EffectiveStatus = .Range("E" & RowIndex).Value2
        Stage = .Range("G" & RowIndex).Value2
        Path = .Range("I" & RowIndex).Value2
        Step = .Range("Y" & RowIndex).Value2
        Approvers = .Range("AA" & RowIndex).Value2
        StepCriteriaDescription = .Range("AD" & RowIndex).Value2
        StepFieldName = .Range("AF" & RowIndex).Value2
        
        Dim Index As Long
        Dim Operator As String
        Operator = .Range("AG" & RowIndex).Value2
        If Operator = "Equals" Then
            Values.Add .Range("AH" & RowIndex).Value2
        ElseIf Operator = "List Values" Then
            '@Ignore UseMeaningfulName
            Dim Str As Variant ' to be used in For Each loop
            Index = 1
            Do While Index <= Departments.Count
                For Each Str In Split(.Range("AH" & RowIndex).Value2, ",")
                    If Str = Departments.Item(Index).DeptID Then AddValue Str
                Next Str
                Index = Index + 1
            Loop
        ' TODO: Either make this work with non-DEPTID criteria or make this sub explicitly for depts only
        ElseIf Operator = "Between" And StepFieldName = "DEPTID" Then
            Dim ValueRange As Variant ' Array([Start], [End])
            ValueRange = Split(.Range("AH" & RowIndex).Value2, " AND ")
            Index = 1
            Do While Index <= Departments.Count
                If Departments.Item(Index).DeptID >= ValueRange(0) And _
                    Departments.Item(Index).DeptID <= ValueRange(1) Then
                    AddValue Departments.Item(Index).DeptID
                End If
                Index = Index + 1
            Loop
        Else
            Debug.Print "Unhandled Step Criteria Operator"
        End If
    End With
End Sub

Public Function IncludesDepartment(ByVal Dept As Department) As Boolean
    Dim Index As Long
    IncludesDepartment = False ' False by default
    For Index = 1 To ValueCount()
        If StepFieldName = "DEPTID" And _
           Value(Index) = Dept.DeptID Then
           IncludesDepartment = True
           Exit Function
        End If
    Next Index
End Function

' TODO: Return security role used for approval routing
