Attribute VB_Name = "C_Cleaning"
Option Explicit

Function Clean_Log() As Boolean
    Procedure = "Clean_Log"
    '********************************************
    '*************** ERROR POLICY ***************
    '********************************************
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    'On Error GoTo Error_Handler
    Err.Number = 0
    '********************************************
    '********** HERE WE GO CLEANING! ************
    '********************************************
    CheckWorksheet ("Log")
    Sheets("Log").Activate
    ActiveWindow.DisplayGridlines = False
    
    With Sheets("Log").Range("A:A")
        .ClearContents
        .ColumnWidth = 150
        .ClearFormats
    End With
    
    Sheets("Log").Range("A1") = "Events log:"
    Sheets("Log").Range("A2") = "Date and Time\Procedure\info or error description:"
    Sheets("Log").Range("A2").Style = "Accent1"
    Sheets("Log").Range("A2").Font.Bold = True
    
    Event_Number = 0
    Sheets("Main").Activate
    '********************************************
    '************** ERROR HANDLING **************
    '********************************************
Error_Handler:
    If Err.Number <> 0 Then
        Call Print_Log(0, Procedure, Err.Description, "Bad")
        Clean_Log = False
    Else
        Clean_Log = True
    End If
    '********************************************
    '************* DEFAULT SETTINGS *************
    '********************************************
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    
End Function

Function Format_Main_Wsh() As Boolean
    Procedure = "Format_Main_Wsh"
    '********************************************
    '*************** ERROR POLICY ***************
    '********************************************
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    On Error GoTo Error_Handler
    Err.Number = 0
    '********************************************
    '******* CLEAN AND REFRESH THE FORMAT *******
    '********************************************
    
    Dim Number_A As Long 'Number of rows to be formatted.
    Dim Temp_Gene As String, Temp_Species As String
    Number_A = 1000
    
    Sheets("Main").Range("A:Z").ClearFormats
    
    Temp_Gene = Sheets("Main").Range("Targeted_Gene")
    Temp_Species = Sheets("Main").Range("Species")
    
    Sheets("Main").Range("A:Z").ClearContents
    
    Sheets("Main").Range("Targeted_Gene").NumberFormat = "@"
    Sheets("Main").Range("Targeted_Gene") = Temp_Gene
    Sheets("Main").Range("Species") = Temp_Species
    
    Sheets("Main").Range(Range("Reference"), Range("CutSite").Offset(Number_A, 0)).ClearContents
    Sheets("Main").Range(Range("sgRNA1"), Range("PAM_Status").Offset(Number_A, 0)).ClearContents
    
    
    'User inputs:
    With Sheets("Main").Range(Range("Targeted_Gene").Offset(0, -1), Range("Species"))
        .Style = "Note"
        .Borders(xlEdgeTop).ColorIndex = 1
        .Borders(xlEdgeTop).Weight = xlMedium
        .Borders(xlEdgeBottom).ColorIndex = 1
        .Borders(xlEdgeBottom).Weight = xlMedium
    End With
    
    Sheets("Main").Range(Range("Targeted_Gene"), Range("Species")).Font.Bold = True
    Sheets("Main").Range("Targeted_Gene").Offset(-1, -1) = "* Required input"
    Sheets("Main").Range("Targeted_Gene").Offset(0, -1) = "* Target Gene:"
    Sheets("Main").Range("Species").Offset(0, -1) = "* Species (Human/Mouse):"
    Sheets("Main").Range("Targeted_Gene").Offset(-1, -1).Font.ColorIndex = 3
    
    Sheets("Main").Range("Reference") = "Reference"
    Sheets("Main").Range("Sequence") = "Sequence"
    Sheets("Main").Range("Annotation_Type") = "Annotation Type"
    Sheets("Main").Range("Annotation_Name") = "Annotation Name"
    Sheets("Main").Range("Strand") = "Strand"
    Sheets("Main").Range("Results") = "Result"
    Sheets("Main").Range("CutSite") = "Cut Site"
    Sheets("Main").Range("sgRNA1") = "sgRNA1"
    Sheets("Main").Range("sgRNA2") = "sgRNA2"
    Sheets("Main").Range("DeletionSize") = "Deletion Size"
    Sheets("Main").Range("Frame") = "Frame"
    Sheets("Main").Range("PAM_Status") = "PAM Status"
    
    'Saved file address:
    Sheets("Main").Range("NewAddress").ClearContents
    Sheets("Main").Range("NewAddress").Offset(-1, 0) = "The updated .gb file can be found here:"
    Sheets("Main").Range("NewAddress").Font.ColorIndex = 5
    Sheets("Main").Range("NewAddress").Font.Bold = True
    
    'Table and Header:
    Sheets("Main").Range(Range("Reference"), Range("CutSite").Offset(Number_A, 0)).Style = "Note"
    Sheets("Main").Range(Range("Annotation_Type"), Range("Annotation_Name").Offset(Number_A, 0)).Style = "Neutral"
    Sheets("Main").Range(Range("Strand"), Range("CutSite").Offset(Number_A, 0)).Style = "Good"
    With Sheets("Main").Range(Range("Reference"), Range("CutSite"))
        .Font.Bold = True
        .Font.Size = 12
        .Borders(xlEdgeTop).ColorIndex = 1
        .Borders(xlEdgeTop).Weight = xlMedium
        .Borders(xlEdgeBottom).ColorIndex = 1
        .Borders(xlEdgeBottom).Weight = xlMedium
    End With

    'Dual-guide combinations:
    Sheets("Main").Range(Range("sgRNA1"), Range("PAM_Status").Offset(Number_A, 0)).Style = "Input"
    With Sheets("Main").Range(Range("sgRNA1"), Range("PAM_Status"))
        .Font.Bold = True
        .Font.Size = 12
        .Borders(xlEdgeTop).ColorIndex = 1
        .Borders(xlEdgeTop).Weight = xlMedium
        .Borders(xlEdgeBottom).ColorIndex = 1
        .Borders(xlEdgeBottom).Weight = xlMedium
    End With
    
    'Distance settings:
    Sheets("Main").Range("Min_Distance_PAM3").Offset(-1, -1) = "Min. distances:"
    Sheets("Main").Range("Min_Distance_PAM3").Offset(0, -1) = "PAM 3« (Fwd-Fwd):"
    Sheets("Main").Range("Min_Distance_PAM5").Offset(0, -1) = "PAM 5« (Rev-Rev):"
    Sheets("Main").Range("Min_Distance_PAMin").Offset(0, -1) = "PAM-in (Fwd-Rev):"
    Sheets("Main").Range("Min_Distance_PAMout").Offset(0, -1) = "PAM-out (Rev-Fwd):"
    Sheets("Main").Range("Max_Distance").Offset(0, -1) = "Max. distance:"
    
    
    If Sheets("Main").Range("Min_Distance_PAM3") < 1 Then Sheets("Main").Range("Min_Distance_PAM3") = 30
    If Sheets("Main").Range("Min_Distance_PAM5") < 1 Then Sheets("Main").Range("Min_Distance_PAM5") = 30
    If Sheets("Main").Range("Min_Distance_PAMin") < 1 Then Sheets("Main").Range("Min_Distance_PAMin") = 20
    If Sheets("Main").Range("Min_Distance_PAMout") < 1 Then Sheets("Main").Range("Min_Distance_PAMout") = 30
    If Sheets("Main").Range("Max_Distance") < 1 Then Sheets("Main").Range("Max_Distance") = 100
    
    Sheets("Main").Range(Range("Min_Distance_PAM3").Offset(-1, -1), Range("Max_Distance")).Style = "Note"
    With Sheets("Main").Range(Range("Min_Distance_PAM3").Offset(-1, -1), Range("Min_Distance_PAM3").Offset(-1, 0))
        .Font.Bold = True
        .Font.Size = 12
        .Borders(xlEdgeTop).ColorIndex = 1
        .Borders(xlEdgeTop).Weight = xlMedium
        .Borders(xlEdgeBottom).ColorIndex = 1
        .Borders(xlEdgeBottom).Weight = xlMedium
    End With
    
    With Sheets("Main").Range(Range("Max_Distance").Offset(0, -1), Range("Max_Distance"))
        .Font.Bold = True
        .Font.Size = 12
        .Borders(xlEdgeTop).ColorIndex = 1
        .Borders(xlEdgeTop).Weight = xlMedium
        .Borders(xlEdgeBottom).ColorIndex = 1
        .Borders(xlEdgeBottom).Weight = xlMedium
    End With
    
    '********************************************
    '************** ERROR HANDLING **************
    '********************************************
Error_Handler:
    If Err.Number <> 0 Then
        Call Print_Log(0, Procedure, Err.Description, "Bad")
        Format_Main_Wsh = False
    Else
        Format_Main_Wsh = True
    End If
    '********************************************
    '************* DEFAULT SETTINGS *************
    '********************************************
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    
    
End Function

Function HouseKeeping() As Boolean
    Procedure = "HouseKeeping"
    '********************************************
    '*************** ERROR POLICY ***************
    '********************************************
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    On Error GoTo Error_Handler
    Err.Number = 0
    '********************************************
    '******* CLEAN AND REFRESH THE FORMAT *******
    '********************************************
    Sheets("Main").Range("NewAddress").ClearContents
    
    Dim Number_A As Long 'Number of rows to be formatted.
    Number_A = 1000
    
    Sheets("Main").Range(Range("Strand"), Range("CutSite").Offset(Number_A, 0)).ClearContents
    Sheets("Main").Range(Range("sgRNA1"), Range("PAM_Status").Offset(Number_A, 0)).ClearContents
    
    'User inputs:
    With Sheets("Main").Range(Range("Targeted_Gene").Offset(0, -1), Range("Species"))
        .Style = "Note"
        .Borders(xlEdgeTop).ColorIndex = 1
        .Borders(xlEdgeTop).Weight = xlMedium
        .Borders(xlEdgeBottom).ColorIndex = 1
        .Borders(xlEdgeBottom).Weight = xlMedium
    End With
    
    Sheets("Main").Range("Targeted_Gene").NumberFormat = "@"
    Sheets("Main").Range(Range("Targeted_Gene"), Range("Species")).Font.Bold = True
    Sheets("Main").Range("Targeted_Gene").Offset(-1, -1) = "* Required input"
    Sheets("Main").Range("Targeted_Gene").Offset(0, -1) = "* Target Gene:"
    Sheets("Main").Range("Species").Offset(0, -1) = "* Species (Human/Mouse):"
    Sheets("Main").Range("Targeted_Gene").Offset(-1, -1).Font.ColorIndex = 3
    
    Sheets("Main").Range("Reference") = "Reference"
    Sheets("Main").Range("Sequence") = "Sequence"
    Sheets("Main").Range("Annotation_Type") = "Annotation Type"
    Sheets("Main").Range("Annotation_Name") = "Annotation Name"
    Sheets("Main").Range("Strand") = "Strand"
    Sheets("Main").Range("Results") = "Result"
    Sheets("Main").Range("CutSite") = "Cut Site"
    Sheets("Main").Range("sgRNA1") = "sgRNA1"
    Sheets("Main").Range("sgRNA2") = "sgRNA2"
    Sheets("Main").Range("DeletionSize") = "Deletion Size"
    Sheets("Main").Range("Frame") = "Frame"
    Sheets("Main").Range("PAM_Status") = "PAM Status"
    
    'Table and Header:
    Sheets("Main").Range(Range("Strand"), Range("CutSite").Offset(Number_A, 0)).Style = "Good"
    With Sheets("Main").Range(Range("Reference"), Range("CutSite"))
        .Font.Bold = True
        .Font.Size = 12
        .Borders(xlEdgeTop).ColorIndex = 1
        .Borders(xlEdgeTop).Weight = xlMedium
        .Borders(xlEdgeBottom).ColorIndex = 1
        .Borders(xlEdgeBottom).Weight = xlMedium
    End With

    'Dual-guide combinations:
    Sheets("Main").Range(Range("sgRNA1"), Range("PAM_Status").Offset(Number_A, 0)).Style = "Input"
    With Sheets("Main").Range(Range("sgRNA1"), Range("PAM_Status"))
        .Font.Bold = True
        .Font.Size = 12
        .Borders(xlEdgeTop).ColorIndex = 1
        .Borders(xlEdgeTop).Weight = xlMedium
        .Borders(xlEdgeBottom).ColorIndex = 1
        .Borders(xlEdgeBottom).Weight = xlMedium
    End With
    
    'Distance settings:
    Sheets("Main").Range("Min_Distance_PAM3").Offset(-1, -1) = "Min. distances:"
    Sheets("Main").Range("Min_Distance_PAM3").Offset(0, -1) = "PAM 3« (Fwd-Fwd):"
    Sheets("Main").Range("Min_Distance_PAM5").Offset(0, -1) = "PAM 5« (Rev-Rev):"
    Sheets("Main").Range("Min_Distance_PAMin").Offset(0, -1) = "PAM-in (Fwd-Rev):"
    Sheets("Main").Range("Min_Distance_PAMout").Offset(0, -1) = "PAM-out (Rev-Fwd):"
    Sheets("Main").Range("Max_Distance").Offset(0, -1) = "Max. distance:"
    
    If Sheets("Main").Range("Min_Distance_PAM3") < 1 Then Sheets("Main").Range("Min_Distance_PAM3") = 30
    If Sheets("Main").Range("Min_Distance_PAM5") < 1 Then Sheets("Main").Range("Min_Distance_PAM5") = 30
    If Sheets("Main").Range("Min_Distance_PAMin") < 1 Then Sheets("Main").Range("Min_Distance_PAMin") = 20
    If Sheets("Main").Range("Min_Distance_PAMout") < 1 Then Sheets("Main").Range("Min_Distance_PAMout") = 30
    If Sheets("Main").Range("Max_Distance") < 1 Then Sheets("Main").Range("Max_Distance") = 100
    
    Sheets("Main").Range(Range("Min_Distance_PAM3").Offset(-1, -1), Range("Max_Distance")).Style = "Note"
    With Sheets("Main").Range(Range("Min_Distance_PAM3").Offset(-1, -1), Range("Min_Distance_PAM3").Offset(-1, 0))
        .Font.Bold = True
        .Font.Size = 12
        .Borders(xlEdgeTop).ColorIndex = 1
        .Borders(xlEdgeTop).Weight = xlMedium
        .Borders(xlEdgeBottom).ColorIndex = 1
        .Borders(xlEdgeBottom).Weight = xlMedium
    End With
    
    With Sheets("Main").Range(Range("Max_Distance").Offset(0, -1), Range("Max_Distance"))
        .Font.Bold = True
        .Font.Size = 12
        .Borders(xlEdgeTop).ColorIndex = 1
        .Borders(xlEdgeTop).Weight = xlMedium
        .Borders(xlEdgeBottom).ColorIndex = 1
        .Borders(xlEdgeBottom).Weight = xlMedium
    End With
    
    '********************************************
    '************** ERROR HANDLING **************
    '********************************************
Error_Handler:
    If Err.Number <> 0 Then
        Call Print_Log(0, Procedure, Err.Description, "Bad")
        HouseKeeping = False
    Else
        HouseKeeping = True
    End If
    '********************************************
    '************* DEFAULT SETTINGS *************
    '********************************************
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    
    
End Function


Function Duals_Housekeeping() As Boolean
    Procedure = "Duals_Housekeeping"
    '********************************************
    '*************** ERROR POLICY ***************
    '********************************************
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    On Error GoTo Error_Handler
    Err.Number = 0
    '********************************************
    '******* CLEAN AND REFRESH THE FORMAT *******
    '********************************************
    
    Dim Number_A As Long 'Number of rows to be formatted.
    Number_A = 1000
    
    Sheets("Main").Range(Range("sgRNA1"), Range("PAM_Status").Offset(Number_A, 0)).ClearContents
    Sheets("Main").Range("Targeted_Gene").NumberFormat = "@"
    
    Sheets("Main").Range("sgRNA1") = "sgRNA1"
    Sheets("Main").Range("sgRNA2") = "sgRNA2"
    Sheets("Main").Range("DeletionSize") = "Deletion Size"
    Sheets("Main").Range("Frame") = "Frame"
    Sheets("Main").Range("PAM_Status") = "PAM Status"

    'Dual-guide combinations:
    Sheets("Main").Range(Range("sgRNA1"), Range("PAM_Status").Offset(Number_A, 0)).Style = "Input"
    With Sheets("Main").Range(Range("sgRNA1"), Range("PAM_Status"))
        .Font.Bold = True
        .Font.Size = 12
        .Borders(xlEdgeTop).ColorIndex = 1
        .Borders(xlEdgeTop).Weight = xlMedium
        .Borders(xlEdgeBottom).ColorIndex = 1
        .Borders(xlEdgeBottom).Weight = xlMedium
    End With
    
    'Distance settings:
    Sheets("Main").Range("Min_Distance_PAM3").Offset(-1, -1) = "Min. distances:"
    Sheets("Main").Range("Min_Distance_PAM3").Offset(0, -1) = "PAM 3´ (Fwd-Fwd):"
    Sheets("Main").Range("Min_Distance_PAM5").Offset(0, -1) = "PAM 5´ (Rev-Rev):"
    Sheets("Main").Range("Min_Distance_PAMin").Offset(0, -1) = "PAM-in (Fwd-Rev):"
    Sheets("Main").Range("Min_Distance_PAMout").Offset(0, -1) = "PAM -out(Rev - Fwd):"
    Sheets("Main").Range("Max_Distance").Offset(0, -1) = "Max. distance:"
    
    If Sheets("Main").Range("Min_Distance_PAM3") < 1 Then Sheets("Main").Range("Min_Distance_PAM3") = 30
    If Sheets("Main").Range("Min_Distance_PAM5") < 1 Then Sheets("Main").Range("Min_Distance_PAM5") = 30
    If Sheets("Main").Range("Min_Distance_PAMin") < 1 Then Sheets("Main").Range("Min_Distance_PAMin") = 20
    If Sheets("Main").Range("Min_Distance_PAMout") < 1 Then Sheets("Main").Range("Min_Distance_PAMout") = 30
    If Sheets("Main").Range("Max_Distance") < 1 Then Sheets("Main").Range("Max_Distance") = 100
    
    Sheets("Main").Range(Range("Min_Distance_PAM3").Offset(-1, -1), Range("Max_Distance")).Style = "Note"
    With Sheets("Main").Range(Range("Min_Distance_PAM3").Offset(-1, -1), Range("Min_Distance_PAM3").Offset(-1, 0))
        .Font.Bold = True
        .Font.Size = 12
        .Borders(xlEdgeTop).ColorIndex = 1
        .Borders(xlEdgeTop).Weight = xlMedium
        .Borders(xlEdgeBottom).ColorIndex = 1
        .Borders(xlEdgeBottom).Weight = xlMedium
    End With
    
    With Sheets("Main").Range(Range("Max_Distance").Offset(0, -1), Range("Max_Distance"))
        .Font.Bold = True
        .Font.Size = 12
        .Borders(xlEdgeTop).ColorIndex = 1
        .Borders(xlEdgeTop).Weight = xlMedium
        .Borders(xlEdgeBottom).ColorIndex = 1
        .Borders(xlEdgeBottom).Weight = xlMedium
    End With
    
    '********************************************
    '************** ERROR HANDLING **************
    '********************************************
Error_Handler:
    If Err.Number <> 0 Then
        Call Print_Log(0, Procedure, Err.Description, "Bad")
        Duals_Housekeeping = False
    Else
        Duals_Housekeeping = True
    End If
    '********************************************
    '************* DEFAULT SETTINGS *************
    '********************************************
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    
    
End Function

Function CheckWorksheet(Wsh As String)
    
    Dim ws As Worksheet
    Err.Number = 0
    
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(Wsh)
    
    If Not Err.Number = 0 Then
        Sheets.Add.Name = Wsh
        Err.Number = 0
    End If
    
End Function

Function Print_Log(i As Long, Procedure As String, Msg As String, Format As String)
    
    On Error Resume Next
    
    Sheets("Log").Range("A3").Offset(Event_Number, 0) = Now & "\" & Procedure & "\" & Str(i) & " \" & Msg
    Sheets("Log").Range("A3").Offset(Event_Number, 0).Style = Format
    Event_Number = Event_Number + 1
    
End Function

