VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ExpenseApprovalCollection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("T&E")

Private Type TExpenseApprovalCollection
    ExpenseApprovals As Collection
End Type
Private this As TExpenseApprovalCollection

' TODO: Implement Key, Before, After
Public Sub Add(ByVal NewItem As ExpenseApproval)
    this.ExpenseApprovals.Add NewItem
End Sub

Public Function Count() As Long
    Count = this.ExpenseApprovals.Count
End Function

Public Function Item(ByVal Index As Variant) As ExpenseApproval
    Item = this.ExpenseApprovals.Item(Index)
End Function

Public Sub Remove(ByVal Index As Variant)
    this.ExpenseApprovals.Remove Index
End Sub
