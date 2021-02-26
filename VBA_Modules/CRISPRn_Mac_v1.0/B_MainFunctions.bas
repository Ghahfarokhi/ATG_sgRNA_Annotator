Attribute VB_Name = "B_MainFunctions"
Option Explicit

Function Import_sgRNAs() As Boolean
    Procedure = "Import_sgRNAs"
    '********************************************
    '*************** ERROR POLICY ***************
    '********************************************
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    On Error GoTo Error_Handler
    Err.Number = 0
    '********************************************
    '*******             MAIN             *******
    '********************************************
    Call Print_Log(0, Procedure, "Importing sgRNAs started!", "Good")
    
    If Load_Coding_Array = True Then
        
        Call Print_Log(0, Procedure, "Coding array is successfully loaded!", "Good")
        
    Else
        
        Call Print_Log(0, Procedure, "Failure in loading the coding array!", "Bad")
        Import_sgRNAs = False
        Exit Function
    End If
    
    Dim Library_Path As String, Gene_Name As String, sgRNA_DataLine As String, i As Long, ParsingLine() As String, Species As String
    
    Gene_Name = Gene_Name_Conversion(Sheets("Main").Range("Targeted_Gene"))
    Sheets("Main").Range("Targeted_Gene") = Gene_Name
    
    If Len(Gene_Name) > 1 Then
        Call Print_Log(0, Procedure, "Selected Gene: " & Gene_Name, "Good")
    Else
        Call Print_Log(0, Procedure, "Please provide a Target Gene!", "Bad")
        Import_sgRNAs = False
        Exit Function
    End If
    
    Species = UCase(Sheets("Main").Range("Species"))
    
    If InStr(1, Species, "HUMAN") > 0 Or InStr(1, Species, "HOMO") > 0 Then
        Library_Path = ActiveWorkbook.Path & "/Library/" & "hCRISPRn_Lib.txt"
        Call Print_Log(0, Procedure, "Selected species: Human.", "Good")
    ElseIf InStr(1, Species, "MOUSE") > 0 Or InStr(1, Species, "MUS") > 0 Then
        Library_Path = ActiveWorkbook.Path & "/Library/" & "mCRISPRn_Lib.txt"
        Call Print_Log(0, Procedure, "Selected species: Mouse.", "Good")
    Else
        Call Print_Log(0, Procedure, "Selected species should be either Human or Mouse.", "Bad")
        Import_sgRNAs = False
        Exit Function
    End If
    
    Temp_Text = Dir(Library_Path)
    
    If Temp_Text = "" Then
        Call Print_Log(0, Procedure, "Couldn't find this library: " & Library_Path, "Bad")
        Import_sgRNAs = False
        Exit Function
    Else
        Call Print_Log(0, Procedure, "Library_Path: " & Library_Path, "Good")
    End If
        
    Gene_Name = ";" + Gene_Name + ";"
    
    i = 1
    
    Open Library_Path For Input As #1
    
    While Not EOF(1)
        Line Input #1, sgRNA_DataLine
        If InStr(1, UCase(sgRNA_DataLine), Gene_Name) > 0 Then
            ParsingLine() = Split(sgRNA_DataLine, ";")
            Range("Reference").Offset(i, 0) = Reference_Identifier(Species, CLng(ParsingLine(0)))
            Range("Sequence").Offset(i, 0) = CODE_2_DNA(ParsingLine(2))
            If InStr(1, ParsingLine(2), "TTTT") > 0 Or Right(ParsingLine(2), 3) = "TTT" Then
                Range("Sequence").Offset(i, 0).Style = "Bad"
            End If
            Range("Annotation_Type").Offset(i, 0) = "CRISPR"
            i = i + 1
        End If
    Wend
    
    Close #1
    
    User_Notification = "Total number of imported sgRNAs: " & Str(i - 1) & " ."
    Call Print_Log(0, Procedure, User_Notification, "Good")

    '********************************************
    '************** ERROR HANDLING **************
    '********************************************
Error_Handler:
    If Err.Number <> 0 Then
        User_Notification = "Error description: " & Err.Description
        On Error Resume Next
        Call Print_Log(0, Procedure, User_Notification, "Bad")
        Import_sgRNAs = False
    Else
        Import_sgRNAs = True
    End If
    '********************************************
    '************* DEFAULT SETTINGS *************
    '********************************************
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    
End Function

Function Count_Records() As Boolean
    Procedure = "Count_Records"
    '********************************************
    '*************** ERROR POLICY ***************
    '********************************************
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    On Error GoTo Error_Handler
    Err.Number = 0
    '********************************************
    '*******             MAIN             *******
    '********************************************
    Total_Records = 0
    
    Dim i As Long
    
    i = Sheets("Main").Range("Sequence").End(xlDown).Row - Sheets("Main").Range("Sequence").Row
    
    If i < 0 Then i = 0
    If i > 100000 Then i = 0
    If i > 10000 Then i = 10000
    
    If i >= 1 Then
        Total_Records = i
        Count_Records = True
        User_Notification = "Total number of records: " & Str(i)
        Call Print_Log(0, Procedure, User_Notification, "Good")
    End If
    '********************************************
    '************** ERROR HANDLING **************
    '********************************************
Error_Handler:
    If Err.Number <> 0 Then
        User_Notification = "Error description: " & Err.Description
        On Error Resume Next
        Call Print_Log(0, Procedure, User_Notification, "Bad")
        Count_Records = False
    End If
    '********************************************
    '************* DEFAULT SETTINGS *************
    '********************************************
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    
End Function

Function Annotator(GenBank_File_Path As String) As Boolean
    Procedure = "Annotator"
    '********************************************
    '*************** ERROR POLICY ***************
    '********************************************
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    On Error GoTo Error_Handler
    Err.Number = 0
    '********************************************
    '*******             MAIN             *******
    '********************************************
    Dim Seq As String, RevCompSeq As String, DataLine As String, i As Long, TotalsgRNAs As Long
    Dim ORIGIN_Found As Boolean
    Dim CRISPR_Features_Added As Boolean
    Dim sgRNA_Strand() As String, sgRNAs() As String, sgRNA_Feature() As String
    Dim AnnotationType() As String, AnnotationName() As String, AnnotationNote() As String
    Dim CutSite() As Long
    Dim Annotated_File_Path As String, Temp_File_Path As String
    
    TotalsgRNAs = Total_Records
    
    Call Print_Log(0, Procedure, "Annotator started!", "Good")
        
    Annotated_File_Path = ActiveWorkbook.Path & "/" & Range("Targeted_Gene") & "_" & "Annotated.gb"
    
    Temp_File_Path = ActiveWorkbook.Path & "/" & "temp.txt"
    
    ORIGIN_Found = False
    CRISPR_Features_Added = False
    
    
    Open GenBank_File_Path For Input As #1
    Open Temp_File_Path For Output As #2
    
    While Not EOF(1)
        Line Input #1, DataLine
        DataLine = Replace(DataLine, Chr(10), vbCrLf)
        Print #2, DataLine
    Wend
    
    Close #1
    Close #2
    
    
    Open Temp_File_Path For Input As #1
    
    While Not EOF(1)
        Line Input #1, DataLine
        If ORIGIN_Found = False And (DataLine = "ORIGIN" Or InStr(1, DataLine, "ORIGIN ") > 0) Then
            ORIGIN_Found = True
            Seq = Mid(DataLine, InStr(1, DataLine, "ORIGIN ") + 7, Len(DataLine))
            GoTo Next_Loop
        ElseIf ORIGIN_Found = True Then
            Seq = Seq + DataLine
        End If
Next_Loop:
    Wend
    Close #1
    
    If ORIGIN_Found = False Then
        User_Notification = "The format of this file: (" & GenBank_File_Path & ") is not GenBank."
        Call Print_Log(0, Procedure, User_Notification, "Bad")
        Annotator = False
        Exit Function
    Else
        Call Print_Log(0, Procedure, "Genbank file passed the Format_Check.", "Good")
    End If
    
    For i = 0 To 9
        Seq = Replace(Seq, i, "")
    Next i
    
    Seq = Replace(Seq, Chr(10), "")
    Seq = Replace(Seq, " ", "")
    Seq = Replace(Seq, "/", "")
    Seq = UCase(Seq)
    
    If Len(Seq) < 5 Then
        User_Notification = "The sequence length of this file: (" & GenBank_File_Path & ") is less than 5 bp."
        Call Print_Log(0, Procedure, User_Notification, "Bad")
        Annotator = False
        Exit Function
    Else
        Call Print_Log(0, Procedure, "Provided sequence length: " & Str(Len(Seq)), "Good")
    End If
    
    RevCompSeq = RevComp(Seq)
    
    ReDim sgRNAs(1 To TotalsgRNAs)
    ReDim AnnotationType(1 To TotalsgRNAs)
    ReDim sgRNA_Strand(1 To TotalsgRNAs)
    ReDim CutSite(1 To TotalsgRNAs)
    ReDim sgRNA_Feature(1 To TotalsgRNAs)
    ReDim AnnotationName(1 To TotalsgRNAs)
    ReDim AnnotationNote(1 To TotalsgRNAs)
    
    For i = 1 To TotalsgRNAs
        sgRNAs(i) = UCase(Range("Sequence").Offset(i, 0))
        If Len(Range("Annotation_Type").Offset(i, 0).Text) = 0 Then
            Range("AnnotationType").Offset(i, 0) = "CRISPR"
            AnnotationType(i) = "CRISPR"
        Else
            AnnotationType(i) = Range("Annotation_Type").Offset(i, 0).Text
        End If
    Next i
    
    For i = 1 To TotalsgRNAs
        If Len(Range("Annotation_Name").Offset(i, 0).Text) = 0 Then
            Range("Annotation_Name").Offset(i, 0) = AnnotationType(i)
            AnnotationName(i) = AnnotationType(i)
        Else
            AnnotationName(i) = Range("Annotation_Name").Offset(i, 0).Text
        End If
    AnnotationNote(i) = Range("Reference").Offset(i, 0)
    Next i
    
    For i = 1 To TotalsgRNAs
        If InStr(1, Seq, sgRNAs(i)) > 0 Then
            sgRNA_Strand(i) = "Fwd"
            CutSite(i) = InStr(1, Seq, sgRNAs(i)) + Len(sgRNAs(i)) - 4
        ElseIf InStr(1, RevCompSeq, sgRNAs(i)) > 0 Then
            sgRNA_Strand(i) = "Rev"
            CutSite(i) = Len(Seq) - InStr(1, RevCompSeq, sgRNAs(i)) - Len(sgRNAs(i)) + 4
        Else
            sgRNA_Strand(i) = "Not found!"
        End If
    Next i
    
    
    For i = 1 To TotalsgRNAs
        Range("Strand").Offset(i, 0) = sgRNA_Strand(i)
        Range("CutSite").Offset(i, 0) = CutSite(i)
    Next i
    
    
    For i = 1 To TotalsgRNAs
        If Not sgRNA_Strand(i) = "Not found!" Then
            If sgRNA_Strand(i) = "Fwd" Then
                sgRNA_Feature(i) = Str(InStr(1, Seq, sgRNAs(i))) & ".." & Str(InStr(1, Seq, sgRNAs(i)) - 1 + Len(sgRNAs(i)))
                sgRNA_Feature(i) = "     " & AnnotationType(i) & " " & Replace(sgRNA_Feature(i), " ", "")
                Range("Results").Offset(i, 0) = "Annotated"
            ElseIf sgRNA_Strand(i) = "Rev" Then
                sgRNA_Feature(i) = Str(InStr(1, Seq, RevComp(sgRNAs(i)))) & ".." & Str(InStr(1, Seq, RevComp(sgRNAs(i))) - 1 + Len(sgRNAs(i)))
                sgRNA_Feature(i) = "     " & AnnotationType(i) & " complement(" & Replace(sgRNA_Feature(i), " ", "") & ")"
                Range("Results").Offset(i, 0) = "Annotated"
            End If
            
            AnnotationName(i) = "     /label=" & AnnotationName(i)
        Else
            Range("Results").Offset(i, 0) = "Not found!"
        End If
    Next i
    
    Open Temp_File_Path For Input As #1
    Open Annotated_File_Path For Output As #2
    
    While Not EOF(1)
        Line Input #1, DataLine
        If CRISPR_Features_Added = False Then
            If Not InStr(1, DataLine, "FEATURES") > 0 Then
                Print #2, DataLine
            Else
                Print #2, DataLine
                For i = 1 To TotalsgRNAs
                    If Not sgRNA_Strand(i) = "Not found!" Then
                        Print #2, sgRNA_Feature(i)
                        Print #2, AnnotationName(i)
                        Print #2, "     /note=" & AnnotationNote(i)
                    End If
                Next i
                CRISPR_Features_Added = True
            End If
        Else
            Print #2, DataLine
        End If
    Wend
    
    Close #1
    Close #2
    
    Range("NewAddress") = Annotated_File_Path
    Call Print_Log(0, Procedure, "Annotated file: " & Annotated_File_Path, "Good")
    
    Call Update_Duals
    
    Kill (Temp_File_Path)
    
    '********************************************
    '************** ERROR HANDLING **************
    '********************************************
Error_Handler:
    If Err.Number <> 0 Then
        User_Notification = "Error description: " & Err.Description
        On Error Resume Next
        Call Print_Log(0, Procedure, User_Notification, "Bad")
        Annotator = False
    Else
        Annotator = True
    End If
    '********************************************
    '************* DEFAULT SETTINGS *************
    '********************************************
    Application.DisplayAlerts = True
End Function

Function Download_File(File_Url As String, Save_Address As String) As Boolean
    Procedure = "Download_File"
    '********************************************
    '*************** ERROR POLICY ***************
    '********************************************
    Application.DisplayAlerts = False
    On Error Resume Next
    Err.Number = 0
    '********************************************
    '*******             MAIN             *******
    '********************************************
    Download_File = False
    
    Set WinHttpReq = CreateObject("Microsoft.XMLHTTP")
    WinHttpReq.Open "GET", File_Url, False
    WinHttpReq.Send
    
    
    If WinHttpReq.Status = 200 Then 'Or WinHttpReq.Status = 0 Then
    
        Set oStream = CreateObject("ADODB.Stream")
        oStream.Open
        oStream.Type = 1
        oStream.Write WinHttpReq.responseBody
        oStream.SaveToFile Save_Address, 2 ' 1 = no overwrite, 2 = overwrite
        oStream.Close
        Download_File = True
        ThisWorkbook.FollowHyperlink File_Url
    
    Else
        If Test_Connection(File_Url, Save_Address) = False Then
            Download_File = False
            Exit Function
        End If
        
    End If
    
    'Check if the saved_File exist here...
    Temp_Text = Dir(Save_Address)
    If Temp_Text = "" Then
        Download_File = False
        Exit Function
    Else
        Call Print_Log(0, Procedure, "Downloading file was successful.", "Good")
        Download_File = True
    End If
    
    '********************************************
    '************** ERROR HANDLING **************
    '********************************************
Error_Handler:
    If Err.Number <> 0 Then
        User_Notification = "Download_File\Error Description: " & Err.Description
        On Error Resume Next
        Call Print_Log(0, Procedure, User_Notification, "Bad")
        Err.Number = 0
        Download_File = False
    Else
        Download_File = True
    End If
    '********************************************
    '************* DEFAULT SETTINGS *************
    '********************************************
    Application.DisplayAlerts = True

End Function


Function Test_Connection(Link As String, Optional Address_To_Save As String) As Boolean
    Procedure = "Test_Connection"
    Test_Connection = False
    
    Err.Number = 0
    On Error Resume Next
    
    Set WinHttpReq = CreateObject("Microsoft.XMLHTTP")
    WinHttpReq.Open "GET", Link, False
    WinHttpReq.Send

    Temp_Counter = WinHttpReq.Status

    If Temp_Counter = 200 Then
        
        Call Print_Log(0, Procedure, "Internet connection is Ok!", "Good")
        Test_Connection = True
        
            If Not Address_To_Save = "" Then
                Set oStream = CreateObject("ADODB.Stream")
                    oStream.Open
                    oStream.Type = 1
                    oStream.Write WinHttpReq.responseBody
                    oStream.SaveToFile Address_To_Save, 2 ' 1 = no overwrite, 2 = overwrite
                    oStream.Close
            End If
        
        Err.Number = 0
        Exit Function
    
    Else
        Call Print_Log(0, Procedure, "Testing the internet connection failed!", "Bad")
        Call Connection_Aid(0, Temp_Counter)
    End If
    
    Err.Number = 0

End Function

Function RevComp(RefSeq As String) As String
    
    On Error Resume Next
    
    Dim RCRefSeq As String
    
    RCRefSeq = Replace(UCase(RefSeq), " ", "")
    RCRefSeq = Replace(Replace(Replace(Replace(Replace(RCRefSeq, "A", "1"), "T", "2"), "C", "3"), "G", "4"), "U", "5")
    RCRefSeq = Replace(Replace(Replace(Replace(Replace(RCRefSeq, "1", "T"), "2", "A"), "3", "G"), "4", "C"), "5", "A")
    RCRefSeq = Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(RCRefSeq, "Y", "1"), "R", "2"), "K", "3"), "M", "4"), "B", "5"), "V", "6"), "D", "7"), "H", "8")
    RCRefSeq = Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(RCRefSeq, "1", "R"), "2", "Y"), "3", "M"), "4", "K"), "5", "V"), "6", "B"), "7", "H"), "8", "D")
    
    RevComp = StrReverse(RCRefSeq)

End Function


Function Defaulter()
    
    Application.DisplayAlerts = True
    Application.DisplayStatusBar = True
    Application.StatusBar = False
    Application.ScreenUpdating = True
    Application.DisplayFormulaBar = True
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic

End Function

Function Connection_Aid(Batch As Long, Req_Status As Long)
    Procedure = "Connection_Aid"
    On Error Resume Next
    
    If Req_Status = 200 Then Call Print_Log(Batch, Procedure, "Internet connection status: OK.", "Good")
    If Req_Status = 100 Then Call Print_Log(Batch, Procedure, "Internet connection status: Continue.", "Bad")
    If Req_Status = 101 Then Call Print_Log(Batch, Procedure, "Internet connection status: Switching protocols.", "Bad")
    
    If Req_Status = 201 Then Call Print_Log(Batch, Procedure, "Internet connection status: Created.", "Bad")
    If Req_Status = 202 Then Call Print_Log(Batch, Procedure, "Internet connection status: Accepted.", "Bad")
    If Req_Status = 203 Then Call Print_Log(Batch, Procedure, "Internet connection status: Non-Authoritative Information.", "Bad")
    If Req_Status = 204 Then Call Print_Log(Batch, Procedure, "Internet connection status: No Content.", "Bad")
    If Req_Status = 205 Then Call Print_Log(Batch, Procedure, "Internet connection status: Reset Content.", "Bad")
    If Req_Status = 206 Then Call Print_Log(Batch, Procedure, "Internet connection status: Partial Content.", "Bad")
    
    If Req_Status = 300 Then Call Print_Log(Batch, Procedure, "Internet connection status: Multiple Choices.", "Bad")
    If Req_Status = 301 Then Call Print_Log(Batch, Procedure, "Internet connection status: Moved Permanently.", "Bad")
    If Req_Status = 302 Then Call Print_Log(Batch, Procedure, "Internet connection status: Found.", "Bad")
    If Req_Status = 303 Then Call Print_Log(Batch, Procedure, "Internet connection status: See Other.", "Bad")
    If Req_Status = 304 Then Call Print_Log(Batch, Procedure, "Internet connection status: Not Modified.", "Bad")
    If Req_Status = 305 Then Call Print_Log(Batch, Procedure, "Internet connection status: Use Proxy.", "Bad")
    If Req_Status = 307 Then Call Print_Log(Batch, Procedure, "Internet connection status: Temporary Redirect.", "Bad")
    
    If Req_Status = 400 Then Call Print_Log(Batch, Procedure, "Internet connection status: Bad Request.", "Bad")
    If Req_Status = 401 Then Call Print_Log(Batch, Procedure, "Internet connection status: Unauthorized.", "Bad")
    If Req_Status = 402 Then Call Print_Log(Batch, Procedure, "Internet connection status: Payment Required.", "Bad")
    If Req_Status = 403 Then Call Print_Log(Batch, Procedure, "Internet connection status: Forbidden.", "Bad")
    If Req_Status = 404 Then Call Print_Log(Batch, Procedure, "Internet connection status: Not Found.", "Bad")
    If Req_Status = 405 Then Call Print_Log(Batch, Procedure, "Internet connection status: Method Not Allowed.", "Bad")
    If Req_Status = 406 Then Call Print_Log(Batch, Procedure, "Internet connection status: Not Acceptable.", "Bad")
    If Req_Status = 407 Then Call Print_Log(Batch, Procedure, "Internet connection status: Proxy Authentication Required.", "Bad")
    If Req_Status = 408 Then Call Print_Log(Batch, Procedure, "Internet connection status: Request Timeout.", "Bad")
    If Req_Status = 409 Then Call Print_Log(Batch, Procedure, "Internet connection status: Conflict.", "Bad")
    If Req_Status = 410 Then Call Print_Log(Batch, Procedure, "Internet connection status: Gone.", "Bad")
    If Req_Status = 411 Then Call Print_Log(Batch, Procedure, "Internet connection status: Length Required.", "Bad")
    If Req_Status = 412 Then Call Print_Log(Batch, Procedure, "Internet connection status: Precondition Failed.", "Bad")
    If Req_Status = 413 Then Call Print_Log(Batch, Procedure, "Internet connection status: Request Entity Too Large.", "Bad")
    If Req_Status = 414 Then Call Print_Log(Batch, Procedure, "Internet connection status: Request-URI Too Long.", "Bad")
    If Req_Status = 415 Then Call Print_Log(Batch, Procedure, "Internet connection status: Unsupported Media Type.", "Bad")
    If Req_Status = 416 Then Call Print_Log(Batch, Procedure, "Internet connection status: Requested Range Not Suitable.", "Bad")
    If Req_Status = 417 Then Call Print_Log(Batch, Procedure, "Internet connection status: Expectation Failed.", "Bad")
    
    If Req_Status = 500 Then Call Print_Log(Batch, Procedure, "Internet connection status: Internal Server Error.", "Bad")
    If Req_Status = 501 Then Call Print_Log(Batch, Procedure, "Internet connection status: Not Implemented.", "Bad")
    If Req_Status = 502 Then Call Print_Log(Batch, Procedure, "Internet connection status: Bad Gateway.", "Bad")
    If Req_Status = 503 Then Call Print_Log(Batch, Procedure, "Internet connection status: Service Unavailable.", "Bad")
    If Req_Status = 504 Then Call Print_Log(Batch, Procedure, "Internet connection status: Gateway Timeout.", "Bad")
    If Req_Status = 505 Then Call Print_Log(Batch, Procedure, "Internet connection status: HTTP Version Not Supported.", "Bad")
    
    Err.Number = 0
    
End Function

Function Update_Duals()
    Procedure = "Update_Duals"
    '********************************************
    '*************** ERROR POLICY ***************
    '********************************************
    Application.DisplayAlerts = False
    On Error Resume Next
    Err.Number = 0
    '********************************************
    '*******             MAIN             *******
    '********************************************
    On Error GoTo Error_Handler
    Dim i As Long, TotalsgRNAs As Long
    Dim sgRNA_Strand() As String, sgRNAs() As String, sgRNA_Feature() As String, AnnotationType() As String
    Dim CutSite() As Long, PAM_Combination() As String, Frame() As String
    Dim Distances() As Long, j As Long, k As Long
    Dim Min_Distance_PAM3 As Long, Min_Distance_PAM5 As Long, Min_Distance_PAMin As Long, Min_Distance_PAMout As Long, Max_Distance As Long
    
    TotalsgRNAs = Total_Records
    
    ReDim sgRNAs(1 To TotalsgRNAs)
    ReDim AnnotationType(1 To TotalsgRNAs)
    ReDim sgRNA_Strand(1 To TotalsgRNAs)
    ReDim CutSite(1 To TotalsgRNAs)
    ReDim sgRNA_Feature(1 To TotalsgRNAs)
    
    For i = 1 To TotalsgRNAs
        sgRNAs(i) = UCase(Range("Sequence").Offset(i, 0))
        sgRNA_Strand(i) = Range("Strand").Offset(i, 0)
        CutSite(i) = Range("CutSite").Offset(i, 0)
    Next i
    
    
    ReDim Distances(1 To TotalsgRNAs, 1 To TotalsgRNAs)
    ReDim PAM_Combination(1 To TotalsgRNAs, 1 To TotalsgRNAs)
    ReDim Frame(1 To TotalsgRNAs, 1 To TotalsgRNAs)
    
    Min_Distance_PAM3 = Range("Min_Distance_PAM3").Value
    Min_Distance_PAM5 = Range("Min_Distance_PAM5").Value
    Min_Distance_PAMin = Range("Min_Distance_PAMin").Value
    Min_Distance_PAMout = Range("Min_Distance_PAMout").Value
    Max_Distance = Range("Max_Distance").Value
    k = 1
    
    For i = 1 To TotalsgRNAs
        For j = 1 To i
            Distances(i, j) = Abs(CutSite(i) - CutSite(j))
            If sgRNA_Strand(i) = "Fwd" And sgRNA_Strand(j) = "Fwd" Then
                PAM_Combination(i, j) = "PAM 3«"
            ElseIf sgRNA_Strand(i) = "Rev" And sgRNA_Strand(j) = "Rev" Then
                PAM_Combination(i, j) = "PAM 5«"
            Else
                If CutSite(i) < CutSite(j) Then
                    If sgRNA_Strand(i) = "Fwd" And sgRNA_Strand(j) = "Rev" Then
                        PAM_Combination(i, j) = "PAM_in"
                    ElseIf sgRNA_Strand(i) = "Rev" And sgRNA_Strand(j) = "Fwd" Then
                        PAM_Combination(i, j) = "PAM_out"
                    End If
                Else
                    If sgRNA_Strand(i) = "Rev" And sgRNA_Strand(j) = "Fwd" Then
                        PAM_Combination(i, j) = "PAM_in"
                    ElseIf sgRNA_Strand(i) = "Fwd" And sgRNA_Strand(j) = "Rev" Then
                        PAM_Combination(i, j) = "PAM_out"
                    End If
                End If
            End If
            
            If (Distances(i, j) Mod 3) = 0 Then
                Frame(i, j) = "inFrame"
            Else
                Frame(i, j) = "FrameShift"
            End If
                
        Next j
    Next i
    
    For i = 1 To TotalsgRNAs
        If Range("Results").Offset(i, 0) = "Not found!" Then
            GoTo next_i
        End If
        For j = 1 To i
            If Range("Results").Offset(i, 0) = "Not found!" Then
                GoTo Next_j
            End If
            If Distances(i, j) < Max_Distance Then
                Select Case PAM_Combination(i, j)
                    Case "PAM_in"
                       If Distances(i, j) > Min_Distance_PAMin Then
                            Range("sgRNA1").Offset(k, 0) = sgRNAs(i)
                            Range("sgRNA2").Offset(k, 0) = sgRNAs(j)
                            Range("DeletionSize").Offset(k, 0) = Distances(i, j)
                            Range("Frame").Offset(k, 0) = Frame(i, j)
                            Range("PAM_Status").Offset(k, 0) = PAM_Combination(i, j)
                            k = k + 1
                       End If
                    Case "PAM_out"
                        If Distances(i, j) > Min_Distance_PAMout Then
                            Range("sgRNA1").Offset(k, 0) = sgRNAs(i)
                            Range("sgRNA2").Offset(k, 0) = sgRNAs(j)
                            Range("DeletionSize").Offset(k, 0) = Distances(i, j)
                            Range("Frame").Offset(k, 0) = Frame(i, j)
                            Range("PAM_Status").Offset(k, 0) = PAM_Combination(i, j)
                            k = k + 1
                        End If
                    Case "PAM 3´"
                        If Distances(i, j) > Min_Distance_PAM3 Then
                            Range("sgRNA1").Offset(k, 0) = sgRNAs(i)
                            Range("sgRNA2").Offset(k, 0) = sgRNAs(j)
                            Range("DeletionSize").Offset(k, 0) = Distances(i, j)
                            Range("Frame").Offset(k, 0) = Frame(i, j)
                            Range("PAM_Status").Offset(k, 0) = PAM_Combination(i, j)
                            k = k + 1
                        End If
                    Case "PAM 5´"
                        If Distances(i, j) > Min_Distance_PAM5 Then
                            Range("sgRNA1").Offset(k, 0) = sgRNAs(i)
                            Range("sgRNA2").Offset(k, 0) = sgRNAs(j)
                            Range("DeletionSize").Offset(k, 0) = Distances(i, j)
                            Range("Frame").Offset(k, 0) = Frame(i, j)
                            Range("PAM_Status").Offset(k, 0) = PAM_Combination(i, j)
                            k = k + 1
                        End If
                End Select
            End If
Next_j:
        Next j
next_i:
    Next i

    '********************************************
    '************** ERROR HANDLING **************
    '********************************************
Error_Handler:
    If Err.Number <> 0 Then
        User_Notification = "Error Description: " & Err.Description
        On Error Resume Next
        Call Print_Log(0, Procedure, User_Notification, "Bad")
    Else
        Call Print_Log(0, Procedure, "Done!", "Good")
    End If
    
End Function


Function Check_Version()
    Procedure = "Check_Version"
    '********************************************
    '*************** ERROR POLICY ***************
    '********************************************
    Application.DisplayAlerts = False
    On Error Resume Next
    Err.Number = 0
    '********************************************
    '*******             MAIN             *******
    '********************************************
    Temp_Text = "https://drive.google.com/uc?export=download&id=1w9pB3b7pL3dGRZGfj9Mv9JqjRbtEtK_6"
    
    Temp_File_Address = ActiveWorkbook.Path & "/Library/" & "Version_sgRNA_Annotator.txt"
    
    Dim Version_Line As String
    Dim New_Version As Long, Current_Version As Long
    
    Current_Version = 1
    
    '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    Set WinHttpReq = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    WinHttpReq.Open "GET", Temp_Text, False
    WinHttpReq.Send
    
    If WinHttpReq.Status = 200 Then
    
        Set oStream = CreateObject("ADODB.Stream")
        oStream.Open
        oStream.Type = 1
        oStream.Write WinHttpReq.responseBody
        oStream.SaveToFile Temp_File_Address, 2 ' 1 = no overwrite, 2 = overwrite
        oStream.Close
        
    End If
    
    '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

    Temp_Text = ""

    Open Temp_File_Address For Input As #1

    While Not EOF(1)

        Line Input #1, Version_Line
        Temp_Text = Temp_Text & Version_Line

    Wend

    Close #1

    Kill Temp_File_Address

    New_Version = CLng(Left(Temp_Text, InStr(1, Temp_Text, "/") - 1))

    If New_Version > Current_Version Then
        Call Print_Log(0, Procedure, "A new version is available. Please download the updated version here: " & _
        "https://github.com/Ghahfarokhi/sgRNA_Annotator", "Note")
    Else
        Call Print_Log(0, Procedure, "RefSeq Downloader is up to date.", "Good")
    End If
    
    '********************************************
    '************** ERROR HANDLING **************
    '********************************************
Error_Handler:
    If Err.Number <> 0 Then
        User_Notification = "Error Description: " & Err.Description
        On Error Resume Next
        Call Print_Log(0, Procedure, User_Notification, "Bad")
    End If
End Function

Function Reference_Identifier(Species As String, Refrence_Number As Long) As String
    
    If UCase(Species) = "HUMAN" Then
        If Refrence_Number = 1 Then Reference_Identifier = "PMID: 28474669 _ Bassik"
        If Refrence_Number = 2 Then Reference_Identifier = "PMID: 31911676 _ Bradley"
        If Refrence_Number = 3 Then Reference_Identifier = "PMID: 26780180 _ Doench"
        If Refrence_Number = 4 Then Reference_Identifier = "PMID: 29720387 _ Lin"
        If Refrence_Number = 5 Then Reference_Identifier = "PMID: 28655737 _ TKOv3"
        If Refrence_Number = 6 Then Reference_Identifier = "PMID: 25961408 _ Vakoc"
        If Refrence_Number = 7 Then Reference_Identifier = "PMID: 29945888 _ Vakoc"
        If Refrence_Number = 8 Then Reference_Identifier = "PMID: 30395134 _ Wei"
        If Refrence_Number = 9 Then Reference_Identifier = "PMID: 25075903 _ GeCKOv2"
    ElseIf UCase(Species) = "MOUSE" Then
        If Refrence_Number = 1 Then Reference_Identifier = "PMID: 28474669 _ Bassik"
        If Refrence_Number = 2 Then Reference_Identifier = "PMID: 29503867 _ Chen"
        If Refrence_Number = 3 Then Reference_Identifier = "PMID: 26780180 _ Doench"
        If Refrence_Number = 4 Then Reference_Identifier = "Pu_Lab: https://www.biorxiv.org/content/10.1101/808402v1"
        If Refrence_Number = 5 Then Reference_Identifier = "PMID: 28162770 _ Sabatini """
        If Refrence_Number = 6 Then Reference_Identifier = "PMID: 30639098 _ Teichmann """
        If Refrence_Number = 7 Then Reference_Identifier = "PMID: 25075903 _ GeCKOv2"
    Else
        Reference_Identifier = Str(Refrence_Number)
    End If
    
End Function

Function Gene_Name_Conversion(Targeted_Gene As String) As String
    
    Targeted_Gene = UCase(Targeted_Gene)
    Gene_Name_Conversion = Targeted_Gene
    
    If Targeted_Gene = "SEPT1" Then Gene_Name_Conversion = "SEPTIN1"
    If Targeted_Gene = "SEPT2" Then Gene_Name_Conversion = "SEPTIN2"
    If Targeted_Gene = "SEPT3" Then Gene_Name_Conversion = "SEPTIN3"
    If Targeted_Gene = "SEPT4" Then Gene_Name_Conversion = "SEPTIN4"
    If Targeted_Gene = "SEPT5" Then Gene_Name_Conversion = "SEPTIN5"
    If Targeted_Gene = "SEPT6" Then Gene_Name_Conversion = "SEPTIN6"
    If Targeted_Gene = "SEPT7" Then Gene_Name_Conversion = "SEPTIN7"
    If Targeted_Gene = "SEPT8" Then Gene_Name_Conversion = "SEPTIN8"
    If Targeted_Gene = "SEPT9" Then Gene_Name_Conversion = "SEPTIN9"
    If Targeted_Gene = "SEPT10" Then Gene_Name_Conversion = "SEPTIN10"
    If Targeted_Gene = "SEPT11" Then Gene_Name_Conversion = "SEPTIN11"
    If Targeted_Gene = "SEPT12" Then Gene_Name_Conversion = "SEPTIN12"
    If Targeted_Gene = "SEPT13" Then Gene_Name_Conversion = "SEPTIN14"

    If Targeted_Gene = "MARCH1" Then Gene_Name_Conversion = "MARCHF1"
    If Targeted_Gene = "MARCH2" Then Gene_Name_Conversion = "MARCHF2"
    If Targeted_Gene = "MARCH3" Then Gene_Name_Conversion = "MARCHF3"
    If Targeted_Gene = "MARCH4" Then Gene_Name_Conversion = "MARCHF4"
    If Targeted_Gene = "MARCH5" Then Gene_Name_Conversion = "MARCHF5"
    If Targeted_Gene = "MARCH6" Then Gene_Name_Conversion = "MARCHF6"
    If Targeted_Gene = "MARCH7" Then Gene_Name_Conversion = "MARCHF7"
    If Targeted_Gene = "MARCH8" Then Gene_Name_Conversion = "MARCHF8"
    If Targeted_Gene = "MARCH9" Then Gene_Name_Conversion = "MARCHF9"
    If Targeted_Gene = "MARCH10" Then Gene_Name_Conversion = "MARCHF10"
    If Targeted_Gene = "MARCH11" Then Gene_Name_Conversion = "MARCHF11"
    
    If Targeted_Gene = "MARC1" Then Gene_Name_Conversion = "MTARC1"
    If Targeted_Gene = "MARC2" Then Gene_Name_Conversion = "MTARC2"
    
    If Targeted_Gene = "DEC1" Then Gene_Name_Conversion = "DELEC1"
    If Targeted_Gene = "OCT4" Then Gene_Name_Conversion = "POU5F1"
    
End Function

Function Load_Coding_Array() As Boolean
    
    On Error Resume Next
    
    Dim Rng As Range
    
    Set Rng = Sheets("Info").Range(Range("DNA_UTF8_Coding").Offset(1, 0), Range("DNA_UTF8_Coding").Offset(84, 2))
    Coding_Array = Rng
    
    If Err.Number = 0 Then
        Load_Coding_Array = True
    Else
        Load_Coding_Array = False
    End If
    
End Function

Function DNA_2_CODE(Seq As String) As String
    
    On Error Resume Next
    
    Dim Coded_Seq As String
    Dim i As Long, j As Long, k As Long
    
    For j = 1 To Len(Seq) \ 3
        
        For k = 1 To 84
            
            If Mid(Seq, (j - 1) * 3 + 1, 3) = Coding_Array(k, 3) Then
                
                Coded_Seq = Coded_Seq & Coding_Array(k, 2)
                GoTo Next_j
                
            End If
            
        Next k
Next_j:
    Next j
        
    If Len(Seq) Mod 3 > 0 Then
        
        For k = 1 To 84
            
            If Right(Seq, Len(Seq) Mod 3) = Coding_Array(k, 3) Then
                
                Coded_Seq = Coded_Seq & Coding_Array(k, 2)
                
            End If
            
        Next k
        
    End If
    
    If Err.Number = 0 Then
        DNA_2_CODE = Coded_Seq
    Else
        DNA_2_CODE = ""
    End If
    
End Function

Function CODE_2_DNA(Coded_Seq As String) As String
    
    On Error Resume Next
    
    Dim Seq As String
    Dim i As Long, j As Long, k As Long
    
    Seq = ""
    
    For j = 1 To Len(Coded_Seq)
        
        For k = 1 To 84
            
            If Mid(Coded_Seq, j, 1) = Coding_Array(k, 2) Then
                
                Seq = Seq & Coding_Array(k, 3)
                GoTo Next_j
                
            End If
            
        Next k
Next_j:
    Next j

    If Err.Number = 0 Then
        CODE_2_DNA = Seq
    Else
        CODE_2_DNA = ""
    End If
    
End Function
