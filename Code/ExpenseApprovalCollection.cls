VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ExpenseApprovalCollection"
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
'@Folder("T&E")

Private Type TExpenseApprovalCollection
    ExpenseApprovals As Collection
End Type
Private this As TExpenseApprovalCollection

Private Sub Class_Initialize()
    Set this.ExpenseApprovals = New Collection
End Sub

' TODO: Implement Key, Before, After
Public Sub Add(ByVal NewItem As ExpenseApproval)
    this.ExpenseApprovals.Add NewItem
End Sub

Public Function Count() As Long
    Count = this.ExpenseApprovals.Count
End Function

Public Function Item(ByVal Index As Variant) As ExpenseApproval
    Set Item = this.ExpenseApprovals.Item(Index)
End Function

Public Sub Remove(ByVal Index As Variant)
    this.ExpenseApprovals.Remove Index
End Sub

Public Sub CreateExpenseApprovalCollectionFromWorksheet(ByVal wsExpenseApprovals As Worksheet)
    '@Ignore UseMeaningfulName
    Dim EA As ExpenseApproval
    Dim RowIndex As Long
    Dim RowCount As Long
    
    RowCount = wsExpenseApprovals.UsedRange.Rows.Count
    RowIndex = 3
    Do While RowIndex <= RowCount
        EA.ReadExpenseApprovalFromWorksheet wsExpenseApprovals, RowIndex
        this.Add EA
        Set EA = Nothing

        RowIndex = RowIndex + 1
    Loop
End Sub

