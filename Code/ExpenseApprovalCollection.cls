VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ExpenseApprovalCollection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
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
    '@Ignore UseMeaningfulName
    Dim rg As Range
    Dim RowIndex As Long
    Dim RowCount As Long
    
    Set rg = wsExpenseApprovals.UsedRange
    RowCount = rg.Rows
    
    RowIndex = 3
    Do While RowIndex <= RowCount
        EA.ReadExpenseApprovalFromWorksheet wsExpenseApprovals, RowIndex
        this.Add EA
        Set EA = Nothing

        RowIndex = RowIndex + 1
    Loop
End Sub

