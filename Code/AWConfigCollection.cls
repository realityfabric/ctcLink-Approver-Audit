VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AWConfigCollection"
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

Private Type TAWConfigCollection
    AWConfigs As Collection ' Collection of ApprovalWorkflowConfig objects
End Type
Private this As TAWConfigCollection

Private Sub Class_Initialize()
    Set this.AWConfigs = New Collection
End Sub

Public Sub Add(ByVal Value As ApprovalWorkflowConfig)
    this.AWConfigs.Add Value
End Sub

Public Sub Remove(ByVal Index As Long)
    this.AWConfigs.Remove Index
End Sub

Public Function Count() As Long
    Count = this.AWConfigs.Count
End Function

Public Function Item(ByVal Index As Long) As ApprovalWorkflowConfig
    Set Item = this.AWConfigs.Item(Index)
End Function

Public Function Filter( _
        Optional ByVal Approvers As String = vbNullString, _
        Optional ByVal DefinitionID As String = vbNullString, _
        Optional ByVal EffectiveDate As Date = #1/1/100#, _
        Optional ByVal EffectiveStatus As String = vbNullString, _
        Optional ByVal StepFieldName As String = vbNullString _
        ) As AWConfigCollection
    Dim Index As Long
    Set Filter = New AWConfigCollection
    
    For Index = 1 To Count
        If (Approvers = vbNullString Or Item(Index).Approvers Like Approvers) And _
           (DefinitionID = vbNullString Or Item(Index).DefinitionID = DefinitionID) And _
           (EffectiveDate = #1/1/100# Or Item(Index).EffectiveDate = EffectiveDate) And _
           (EffectiveStatus = vbNullString Or Item(Index).EffectiveStatus = EffectiveStatus) And _
           (StepFieldName = vbNullString Or Item(Index).StepFieldName = StepFieldName) _
           Then
           Filter.Add Item(Index)
        Else
            Debug.Print "Else"
        End If
    Next Index
End Function
