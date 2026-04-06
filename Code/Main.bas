Attribute VB_Name = "Main"
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

'@Ignore EncapsulatePublicField
Public Sesh As Session

'@VariableDescription("Stores the Unix Timestamp at runtime, set in the Main method.")
Private UnixTimestamp As LongLong
Attribute UnixTimestamp.VB_VarDescription = "Stores the Unix Timestamp at runtime, set in the Main method."

'@Description("Returns the Unix Timestamp recorded at runtime in the Main method.")
Public Function GetTimestamp() As LongLong
Attribute GetTimestamp.VB_Description = "Returns the Unix Timestamp recorded at runtime in the Main method."
    GetTimestamp = UnixTimestamp
End Function

'@Description("Returns the Unix Timestamp recorded at runtime in the Main method as a string.")
Public Function GetTimestampStr() As String
Attribute GetTimestampStr.VB_Description = "Returns the Unix Timestamp recorded at runtime in the Main method as a string."
    GetTimestampStr = Trim$(Str$(GetTimestamp()))
End Function

Public Function UnixTime() As LongLong
    UnixTime = DateDiff("s", "1/1/1970 00:00:00", Now)
End Function

'@EntryPoint
Public Sub Main()
    Set Sesh = New Session
    UnixTimestamp = UnixTime()

    Dim wbOutput As Workbook

    Dim wbApprovalSetup As Workbook
    Dim wbDepartments As Workbook
    Dim wbExpenseApprovers As Workbook
    Dim wbSecurityRoles As Workbook

    '@Ignore UseMeaningfulName
    Dim ws As Worksheet
    Dim wsApprovalSetup As Worksheet
    Dim wsDepartments As Worksheet
    Dim wsExpenseApprovers As Worksheet
    Dim wsSecurityRoles As Worksheet
    Dim wsApproverRolesOverview As Worksheet
    Dim wsDepartmentsOverview As Worksheet

    Dim ExpenseApprovals As ExpenseApprovalCollection
    Dim Departments As DepartmentCollection
    Set ExpenseApprovals = New ExpenseApprovalCollection
    Set Departments = New DepartmentCollection

    ' Show file selection form
    FileSelection.Show
    ' if file selection form was closed without clicking the button to run this application then terminate
    If Sesh.FormClosedWithoutRunning Then Exit Sub

    Set wbOutput = Workbooks.Add
    With wbOutput
        .SaveAs Filename:="ApproverValidation_" & GetTimestampStr()
    End With

    ' Copy Approval Setup into Output Workbook
    Set wbApprovalSetup = Workbooks.Open( _
        Filename:=Sesh.fApprovalSetup _
        , ReadOnly:=True _
    )

    Set ws = wbApprovalSetup.Sheets.Item(1)

    With wbOutput
        ws.Copy After:=.Sheets.Item(.Sheets.Count)
        Set wsApprovalSetup = .Sheets.Item(.Sheets.Count)
        wsApprovalSetup.Name = "Approval Setup"
    End With
    Set ws = Nothing
    wbApprovalSetup.Close

    ' Copy Departments into Output Workbook
    Set wbDepartments = Workbooks.Open( _
        Filename:=Sesh.fDepartments _
        , ReadOnly:=True _
    )

    Set ws = wbDepartments.Sheets.Item(1)

    With wbOutput
        ws.Copy After:=.Sheets.Item(.Sheets.Count)
        Set wsDepartments = .Sheets.Item(.Sheets.Count)
        wsDepartments.Name = "Departments"
    End With
    Set ws = Nothing
    wbDepartments.Close

    ' Copy Expense Approvers into Output Workbook
    Set wbExpenseApprovers = Workbooks.Open( _
        Filename:=Sesh.fExpenseApprovers _
        , ReadOnly:=True _
    )

    Set ws = wbExpenseApprovers.Sheets.Item(1)

    With wbOutput
        ws.Copy After:=.Sheets.Item(.Sheets.Count)
        Set wsExpenseApprovers = .Sheets.Item(.Sheets.Count)
        wsExpenseApprovers.Name = "Expense Approvers"
    End With
    Set ws = Nothing
    wbExpenseApprovers.Close

    ' Copy User Roles into Output Workbook
    Set wbSecurityRoles = Workbooks.Open( _
        Filename:=Sesh.fUserRoles _
        , ReadOnly:=True _
    )
    Set ws = wbSecurityRoles.Sheets.Item(1)
    With wbOutput
        ws.Copy After:=.Sheets.Item(.Sheets.Count)
        Set wsSecurityRoles = .Sheets.Item(.Sheets.Count)
        wsSecurityRoles.Name = "User Roles"
    End With
    Set ws = Nothing
    wbSecurityRoles.Close

    ' Prepare the Approver Roles Overview sheet
    Set wsApproverRolesOverview = wbOutput.Sheets.Item("sheet1")
    With wsApproverRolesOverview
        ' Set Approver Roles Overview sheet name and switch to that sheet.
        .Name = "Roles Overview"
        .Activate

        ' Define the headers
        Dim headerArray As Variant
        headerArray = Array( _
            "EmplID", _
            "Name", _
            "HR Status", _
            "Department Manager", _
            "Travel Approver", _
            "PU/AP Approver", _
            "ZZ Purchasing Approval", _
            "ZZ Requisition Approval", _
            "ZZ Voucher Approval", _
            "ZZ_AW_AP_REVIEW", _
            "ZZ_AW_BI_INV", _
            "ZZ_AW_GRANT_COORDINATOR", _
            "ZZ_AW_PURCHASING_REVIEW", _
            "ZZ_AW_AMT_LEVEL_X", _
            "ZZ_AW_COMMODITY_X", _
            "ZZ_AW_EXEC_LEVEL_X", _
            "ISSUE DETECTED")

        ' Apply the headers to Row 1, then make it pretty.
        .Range("A1").Resize(ColumnSize:=UBound(headerArray) + 1).Value2 = headerArray
        .Range("C1:P1").Orientation = xlDownward
        .Range("C1:P1").VerticalAlignment = xlVAlignCenter

        ' Copy EmplIDs and Names, then HR Status, then make it pretty.
        wsSecurityRoles.Range("C3:D20000").Copy .Range("A2")
        wsSecurityRoles.Range("K3:K20000").Copy .Range("C2")
        .Range("A:Q").Columns.AutoFit

        ' Remove duplicate EmplID, Name, HR Status combos.
        .Range("A:C").RemoveDuplicates Columns:=Array(1, 2, 3), Header:=xlYes

        ' If an employee has the security role listed in Row 1 then show an X.
        .Range("G2:M2000").Formula = _
            "=IF(COUNTIFS('User Roles'!$C:$C,$A2,'User Roles'!$G:$G,G$1)>0,""X"","""")"
        ' If an employee has a security role beginning with the text in Row 1 (minus the trailng "_X", show an X.
        .Range("N2:P2000").Formula = _
            "=IF(COUNTIFS('User Roles'!$C:$C,$A2,'User Roles'!$G:$G,LEFT(N$1,LEN(N$1) - 2) & ""*"")>0,""X"","""")"

        ' If the employee has any approver roles, show X
        .Range("F2:F2000").Formula = _
            "=IF(COUNTIF($G2:$P2,""X"") > 0,""X"", """")"

        ' If the employee is a department manager, show X
        .Range("D2:D2000").Formula = _
            "=IF(COUNTIFS(Departments!H:H,$A2) > 0,""X"","""")"

        ' If the employee is an expense approver, show X
        .Range("E2:E2000").Formula = _
            "=IF(COUNTIFS('Expense Approvers'!C:C,$A2) > 0, ""X"","""")"

        ' Replace formulas with plain text.
        ' We definitely do not want thousands of calculations happening needlessly after the initial run.
        .UsedRange.Value2 = .UsedRange.Value2

        ' Issue Detection
        ' Issues should be indicated textually and NOT only using conditional formatting
        ' Color-based indicators are sometimes inaccessible to those with colorblindness.
        ' Conditional formatting using color is acceptable in addition to non-color visual indicators.

        ' Is the employee inactive with approval roles?
        .Range("R2:R2000").Formula = _
            "=IF(AND($C2 =""I"", $F2 = ""X""), ""Inactive Employee with Approval Roles!"", """")"

        ' Is the employee inactive and a Department Manager?
        .Range("S2:S2000").Formula = _
            "=IF(AND($C2 =""I"", $D2 = ""X""), ""Inactive Employee is Department Manager!"", """")"

        ' Is the employee inactive and an Expense Approver?
        .Range("T2:T2000").Formula = _
            "=IF(AND($C2 =""I"", $E2 = ""X""), ""Inactive Employee is Expense Approver!"", """")"

        ' Is the employee inactive and has AWE routing roles?
        .Range("U2:U2000").Formula = _
            "=IF(AND($C2 =""I"", COUNTA($J2:$P2) > 0), ""Inactive Employee has AWE Routing Roles!"", """")"
        
        ' Is the employee an Expense Approver but not a Department Manager?
        ' TODO: Exclude approvers who have non EXAPPROVER ExpenseApproval roles.
        .Range("V2:V2000").Formula = _
            "=IF(AND($D2 <> ""X"", $E2 = ""X""),""Travel Approver is not Department Manager! (Do they have non-Department approval responsibilities?)"", """")"

        ' Is the employee able to approve everything they need to approve? i.e. do they have Req, PO, and Voucher approval roles?
        .Range("W2:W2000").Formula = _
            "=IF(AND(OR(D2=""X"",J2=""X"",K2=""X"",L2=""X"",N2=""X"",O2=""X"",P2=""X""),OR(NOT(G2=""X""),NOT(H2=""X""),NOT(I2=""X""))), ""Employee is missing an approval role!"","""")"

        .Range("X2:X2000").Formula = _
            "=IF(AND(M2=""X"", NOT(G2 = ""X"")), ""Employee is missing an approval role!"", """")"

        ' Combine issue checks into a single cell, convert to plain text
        '   then delete extraneous
        .Range("Q2:Q2000").Formula = _
            "=TRIM(CONCAT($R2, "" "", $S2,"" "", $T2, "" "", $U2, "" "", $V2, "" "", $W2, "" "", $X2))"
        .Range("Q2:Q2000").Value2 = _
            .Range("Q2:Q2000").Value2
        .Range("R:X") _
            .Columns.Delete
    End With
    
    ' Prepare the Departments Overview sheet
    Set wsDepartmentsOverview = wbOutput.Sheets.Add(After:=wsApproverRolesOverview)
    With wsDepartmentsOverview
        .Name = "Departments Overview"
        .Activate

        ' Department Manager not assigned to Expense Approvals for relevant departments
        ExpenseApprovals.CreateExpenseApprovalCollectionFromWorksheet wsExpenseApprovers
        Departments.AddDepartmentsFromWorksheet wsDepartments

        Dim MismatchedDepartments As DepartmentCollection
        Set MismatchedDepartments = Departments.DepartmentsWithExpenseApproverMismatch(ExpenseApprovals)
        
        .Range("A1:D1").Value2 = Array( _
            "DeptID", _
            "Description", _
            "ManagerID", _
            "Issues" _
        )
        
        Dim Index As Long
        Dim RowIndex As Long
        Index = 1
        RowIndex = 2
        Do While Index <= MismatchedDepartments.Count
            .Range("A" & RowIndex & ":D" & RowIndex).Value2 = Array( _
                MismatchedDepartments.Item(Index).DeptID, _
                MismatchedDepartments.Item(Index).Description, _
                MismatchedDepartments.Item(Index).ManagerID, _
                "Expense Approver does not match!" _
            )
            RowIndex = RowIndex + 1
            Index = Index + 1
        Loop
        
        Dim NoExApprDepartments As DepartmentCollection
        Set NoExApprDepartments = Departments.DepartmentsWithoutExpenseApproval(ExpenseApprovals)
        
        Index = 1
        Do While Index <= NoExApprDepartments.Count
            .Range("A" & RowIndex & ":D" & RowIndex).Value2 = Array( _
                NoExApprDepartments.Item(Index).DeptID, _
                NoExApprDepartments.Item(Index).Description, _
                NoExApprDepartments.Item(Index).ManagerID, _
                "Department has no ExApprover!" _
            )
            RowIndex = RowIndex + 1
            Index = Index + 1
        Loop
    End With
End Sub

