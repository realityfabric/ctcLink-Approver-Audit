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

Public Property Get Value(ByVal Index As Long) As String
    Value = this.Values.Item(Index)
End Property

Public Sub AddValue(ByVal Value As String)
    this.Values.Add Value
End Sub

Public Sub RemoveValue(ByVal Index As Long)
    this.Values.Remove (Index)
End Sub

Public Function ValueCount() As Long
    ValueCount = this.Values.Count
End Function
