Attribute VB_Name = "TestModule_EACollection"
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
'@Folder("Tests.T&E")

Option Explicit
Option Private Module

Private Assert As Object
Private Fakes As Object

'@Ignore UseMeaningfulName
Private EA1 As ExpenseApproval
'@Ignore UseMeaningfulName
Private EA2 As ExpenseApproval
'@Ignore UseMeaningfulName
Private EA3 As ExpenseApproval
'@Ignore UseMeaningfulName
Private EA4 As ExpenseApproval
'@Ignore UseMeaningfulName
Private EA5 As ExpenseApproval
'@Ignore UseMeaningfulName

'@ModuleInitialize
Private Sub ModuleInitialize()
    'this method runs once per module.
    Set Assert = CreateObject("Rubberduck.AssertClass")
    Set Fakes = CreateObject("Rubberduck.FakesProvider")
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
    'this method runs once per module.
    Set Assert = Nothing
    Set Fakes = Nothing
End Sub

'@TestInitialize
Private Sub TestInitialize()
    'This method runs before every test in the module..
    Set EA1 = New ExpenseApproval
    Set EA2 = New ExpenseApproval
    Set EA3 = New ExpenseApproval
    Set EA4 = New ExpenseApproval
    Set EA5 = New ExpenseApproval
    
    EA1.ApproverType = "EXAPPROVER"
    EA1.BusinessUnit = "WA190"
    EA1.DeptDesc = vbNullString
    EA1.EmplID = "1"
    EA1.FirstName = "Karl"
    EA1.LastName = "Marx"
    EA1.FromChartfield = "00001"
    EA1.ToChartfield = "00001"
    
    EA2.ApproverType = "EXAPPROVER"
    EA2.BusinessUnit = "WA190"
    EA2.DeptDesc = vbNullString
    EA2.EmplID = "2"
    EA2.FirstName = "Adam"
    EA2.LastName = "Smith"
    EA2.FromChartfield = "00002"
    EA2.ToChartfield = "00005"
    
    EA3.ApproverType = "EXAPPROVER"
    EA3.BusinessUnit = "WA190"
    EA3.DeptDesc = vbNullString
    EA3.EmplID = "3"
    EA4.FirstName = "John"
    EA5.LastName = "Keynes"
    EA5.FromChartfield = "CNV19"
    EA5.ToChartfield = "CNV19"
    
    EA4.ApproverType = "EXAPPROVER"
    EA4.BusinessUnit = "WA190"
    EA4.DeptDesc = vbNullString
    EA4.EmplID = "4"
    EA4.FirstName = "Ludwig"
    EA4.LastName = "Mises"
    EA4.FromChartfield = "00111"
    EA4.ToChartfield = "10101"
    
    EA5.ApproverType = "EXAPPROVER"
    EA5.BusinessUnit = "WA190"
    EA5.DeptDesc = vbNullString
    EA5.EmplID = "5"
    EA5.FirstName = "John"
    EA5.LastName = "Mill"
    EA5.FromChartfield = "20000"
    EA5.ToChartfield = "20001"
End Sub

'@TestCleanup
Private Sub TestCleanup()
    'this method runs after every test in the module.
    Set EA1 = Nothing
    Set EA2 = Nothing
    Set EA3 = Nothing
    Set EA4 = Nothing
    Set EA5 = Nothing
End Sub

'@TestMethod("No Fail")
Private Sub TestMethod_EACAdd_NoFail()
    On Error GoTo TestFail

    'Arrange:
    Dim EAC As ExpenseApprovalCollection
    Set EAC = New ExpenseApprovalCollection

    'Act:
    EAC.Add EA1
    EAC.Add EA2
    EAC.Add EA3
    EAC.Add EA4
    EAC.Add EA5

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


'@TestMethod("Filter")
Private Sub TestMethod_FilterByEmplID_CorrectCount()
    On Error GoTo TestFail
    
    'Arrange:
    '@Ignore UseMeaningfulName
    Dim EAC As ExpenseApprovalCollection
    '@Ignore UseMeaningfulName
    Dim EACFiltered As ExpenseApprovalCollection
    Set EAC = New ExpenseApprovalCollection
    EAC.Add EA1
    EAC.Add EA2
    EAC.Add EA3
    EAC.Add EA4
    EAC.Add EA5
    
    'Act:
    Set EACFiltered = EAC.Filter(EmplID:="1")
    
    'Assert:
    Assert.IsTrue EACFiltered.Count = 1

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Filter")
Private Sub TestMethod_FilterByEmplID_CorrectEmplID()
    On Error GoTo TestFail
    
    'Arrange:
    '@Ignore UseMeaningfulName
    Dim EAC As ExpenseApprovalCollection
    '@Ignore UseMeaningfulName
    Dim EACFiltered As ExpenseApprovalCollection
    Set EAC = New ExpenseApprovalCollection
    EAC.Add EA1
    EAC.Add EA2
    EAC.Add EA3
    EAC.Add EA4
    EAC.Add EA5
    
    'Act:
    Set EACFiltered = EAC.Filter(EmplID:="1")
    
    'Assert:
    Assert.IsTrue EACFiltered.Item(1).EmplID = "1"

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Filter")
Private Sub TestMethod_FilterNone_CorrectCount()
    On Error GoTo TestFail
    
    'Arrange:
    '@Ignore UseMeaningfulName
    Dim EAC As ExpenseApprovalCollection
    '@Ignore UseMeaningfulName
    Dim EACFiltered As ExpenseApprovalCollection
    Set EAC = New ExpenseApprovalCollection
    EAC.Add EA1
    EAC.Add EA2
    EAC.Add EA3
    EAC.Add EA4
    EAC.Add EA5
    
    'Act:
    Set EACFiltered = EAC.Filter()
    
    'Assert:
    Assert.IsTrue EAC.Count = EACFiltered.Count

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub
