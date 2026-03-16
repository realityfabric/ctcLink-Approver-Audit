VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ExpenseApproval"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("ApproverValidation")
Option Explicit

Private Type TExpenseApproval
    BusinessUnit As String
    ApproverType As String
    EmplID As String
    DeptDesc As String
    FromChartfield As String
    ToChartfield As String
    LastName As String
    FirstName As String
End Type
Private this As TExpenseApproval

Public Property Let BusinessUnit(ByVal Value As String)
    this.BusinessUnit = Value
End Property

Public Property Get BusinessUnit() As String
    BusinessUnit = this.BusinessUnit
End Property

Public Property Let ApproverType(ByVal Value As String)
    this.ApproverType = Value
End Property

Public Property Get ApproverType() As String
    ApproverType = this.ApproverType
End Property

Public Property Let EmplID(ByVal Value As String)
    this.EmplID = Value
End Property

Public Property Get EmplID() As String
    EmplID = this.EmplID
End Property

Public Property Let DeptDesc(ByVal Value As String)
    this.DeptDesc = Value
End Property

Public Property Get DeptDesc() As String
    DeptDesc = this.DeptDesc
End Property

Public Property Let FromChartfield(ByVal Value As String)
    this.FromChartfield = Value
End Property

Public Property Get FromChartfield() As String
    FromChartfield = this.FromChartfield
End Property

Public Property Let ToChartfield(ByVal Value As String)
    this.ToChartfield = Value
End Property

Public Property Get ToChartfield() As String
    ToChartfield = this.ToChartfield
End Property

Public Property Let LastName(ByVal Value As String)
    this.LastName = Value
End Property

Public Property Get LastName() As String
    LastName = this.LastName
End Property

Public Property Let FirstName(ByVal Value As String)
    this.FirstName = Value
End Property

Public Property Get FirstName() As String
    FirstName = this.FirstName
End Property

Public Sub ReadExpenseApprovalFromWorksheet(ByVal ws As Worksheet, ByVal Row As Long)
    Dim DataArray As Variant
    DataArray = ws.Range("A" & Row & ":H" & Row).Value2
    
    BusinessUnit = DataArray(1, 1)
    ApproverType = DataArray(1, 2)
    EmplID = DataArray(1, 3)
    DeptDesc = DataArray(1, 4)
    FromChartfield = DataArray(1, 5)
    ToChartfield = DataArray(1, 6)
    LastName = DataArray(1, 7)
    FirstName = DataArray(1, 8)
End Sub
