Attribute VB_Name = "TestModule_ExpenseApproval"
'@TestModule
'@Folder("Tests")


Option Explicit
Option Private Module

Private Assert As Object
Private Fakes As Object

Private TestWorksheet As Worksheet

'@ModuleInitialize
Private Sub ModuleInitialize()
    'this method runs once per module.
    Set Assert = CreateObject("Rubberduck.AssertClass")
    Set Fakes = CreateObject("Rubberduck.FakesProvider")

    Set TestWorksheet = ThisWorkbook.Sheets.Add
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
    'this method runs once per module.
    Dim DisplayAlerts As Boolean
    
    Set Assert = Nothing
    Set Fakes = Nothing
    
    ' Save Application.DisplayAlerts value, set to False, delete the test sheet,
    '  then set Application.DisplayAlerts back to its previous value.
    DisplayAlerts = Application.DisplayAlerts
    Application.DisplayAlerts = False
    TestWorksheet.Delete
    Application.DisplayAlerts = DisplayAlerts
End Sub

'@TestInitialize
Private Sub TestInitialize()
    With TestWorksheet
        .Range("A1:H1").Value2 = Array( _
            "GL Unit", _
            "Approver Type", _
            "EmplID", _
            "Description", _
            "From Chartfield", _
            "To Chartfield", _
            "Last Name", _
            "First Name" _
        )
        
        .Range("A2:H2").Value2 = Array( _
            "WA190", _
            "EXAPPROVER", _
            "111111111", _
            "Test Department", _
            "10000", _
            "10500", _
            "John", _
            "Smith" _
        )
    End With
End Sub

'@TestCleanup
Private Sub TestCleanup()
    With TestWorksheet
        .UsedRange.Clear
    End With
End Sub

'@TestMethod("Doesn't Crash")
Private Sub TestMethod_ReadExpenseApprovalFromWorksheet_NoCrash()
    On Error GoTo TestFail
    
    'Arrange:
    Dim EA As ExpenseApproval
    Set EA = New ExpenseApproval
    
    'Act:
    EA.ReadExpenseApprovalFromWorksheet TestWorksheet, 2
    
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
