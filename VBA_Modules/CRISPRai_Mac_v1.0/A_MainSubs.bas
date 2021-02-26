Attribute VB_Name = "A_Main_Subs"
    '#############################################
    '############# <  INFORMATION  > #############
    '#############################################
    '###                                       ###
    '###          CRISPRai Annotator           ###
    '###              Version 1.0              ###
    '###              2020 May 30              ###
    '###                                       ###
    '###                                       ###
    '###               Author:                 ###
    '###        Amir Taheri Ghahfarokhi        ###
    '###                                       ###
    '###               Email:                  ###
    '###   Amir.Taheri.Ghahfarokhi@Gmail.com   ###
    '###                                       ###
    '###               GitHub                  ###
    '###    https://github.com/Ghahfarokhi/    ###
    '###                                       ###
    '###                                       ###
    '#############################################
    '#############################################
    '#############################################

'=======================================================
Option Explicit
'Variable declaration:
Public Const Tool_Name = "CRISPRai Annotator v1.0"
Public Total_Records As Long
Public Event_Number As Long
Public User_Notification As String
Public Procedure As String
Public Temp_Text As String
Public Temp_Counter As Long
Public Temp_File_Address As String
Public Coding_Array() As Variant
Public oStream As Object
Public WinHttpReq As Object

Function Reset_Workbook(Optional Mode As String) As Boolean
    
    On Error Resume Next
    
    If Clean_Log = True Then
        If Format_Main_Wsh = True Then
            If Not Mode = "Silent" Then
                MsgBox "Workbook_reset is done!", vbInformation, Tool_Name
                Reset_Workbook = True
            End If
        Else
            If Not Mode = "Silent" Then MsgBox "Reseting the Main worksheet has been failed!", vbInformation, Tool_Name
            Reset_Workbook = False
        End If
    Else
        If Not Mode = "Silent" Then MsgBox "Reseting the Log worksheet has been failed!", vbInformation, Tool_Name
        Reset_Workbook = False
    End If
    
    Call Defaulter
    
End Function

Sub Import_Litrature_sgRNAs()
    
    On Error GoTo Error_Handler
    
    Call Reset_Workbook("Silent")
    'Call Check_Version
    Call Print_Log(0, "Import_Litrature_sgRNAs", "Importing the litrature sgRNAs...", "Good")
    
    If Len(Range("Targeted_Gene")) > 1 Then
        If Import_sgRNAs = True Then
            MsgBox "Import is complete.", vbInformation, Tool_Name
        Else
            MsgBox "Importing sgRNAs failed! Please check the Log worksheet.", vbInformation, Tool_Name
        End If
    Else
        MsgBox "Please provide a gene name!", vbInformation, Tool_Name
    End If
    
Error_Handler:
    If Err.Number <> 0 Then
        User_Notification = Err.Description
        On Error Resume Next
        Call Print_Log(0, "Import_Literature_sgRNAs", User_Notification, "Bad")
        MsgBox "Importing sgRNAs failed! Please check the Log worksheet.", vbInformation, Tool_Name
    Else
        Call Print_Log(0, "Import_Literature_sgRNAs", "Done!", "Good")
    End If
    
    Call Defaulter
    
End Sub

Sub Annotate_sgRNAs_Browse()
    
    On Error GoTo Error_Handler

    Call HouseKeeping
    Call Clean_Log
    'Call Check_Version
    Call Print_Log(0, "Annotate_sgRNAs_Browse", "Annotate_sgRNAs_Browse Started!", "Good")
    
    If Count_Records = True Then
        Dim UserChoice As Integer
        Dim GenBank_File_Path As String

        
        Dim MyPath As String
        Dim MyScript As String
        Dim MyFiles As String
        Dim MySplit As Variant
        Dim N As Long
        Dim Fname As String
        Dim mybook As Workbook
    
        On Error Resume Next
        MyPath = MacScript("return (path to documents folder) as String")
        'Or use MyPath = "Macintosh HD:Users:Ron:Desktop:TestFolder:"
    
        ' In the following statement, change true to false in the line "multiple
        ' selections allowed true" if you do not want to be able to select more
        ' than one file. Additionally, if you want to filter for multiple files, change
        ' {""gb""} to
        ' {""com.microsoft.excel.xls"",""public.comma-separated-values-text""}
        ' if you want to filter on xls and csv files, for example.
        MyScript = _
        "set applescript's text item delimiters to "","" " & vbNewLine & _
                   "set theFiles to (choose file of type " & _
                 " {""gb""} " & _
                   "with prompt ""Please select a .gb file"" default location alias """ & _
                   MyPath & """ multiple selections allowed false) as string" & vbNewLine & _
                   "set applescript's text item delimiters to """" " & vbNewLine & _
                   "return theFiles"
    
        MyFiles = MacScript(MyScript)
'            On Error GoTo 0
'
'            If MyFiles <> "" Then
'                With Application
'                    .ScreenUpdating = False
'                    .EnableEvents = False
'                End With
'
'                MySplit = Split(MyFiles, ",")
'                For N = LBound(MySplit) To UBound(MySplit)
'
'                     Get the file name only and test to see if it is open.
'                    Fname = Right(MySplit(N), Len(MySplit(N)) - InStrRev(MySplit(N), Application.PathSeparator, , 1))
'                    If bIsBookOpen(Fname) = False Then
'
'                        Set mybook = Nothing
'                        On Error Resume Next
'                        Set mybook = Workbooks.Open(MySplit(N))
'                        On Error GoTo 0
'
'                        If Not mybook Is Nothing Then
'                            MsgBox "You open this file : " & MySplit(N) & vbNewLine & _
'                                   "And after you press OK it will be closed" & vbNewLine & _
'                                   "without saving, replace this line with your own code."
'                            mybook.Close SaveChanges:=False
'                        End If
'                    Else
'                        MsgBox "We skipped this file : " & MySplit(N) & " because it Is already open."
'                    End If
'                Next N
'                With Application
'                    .ScreenUpdating = True
'                    .EnableEvents = True
'                End With
'            End If


        MyFiles = Replace(MyFiles, ":", "/")
        MyFiles = Right(MyFiles, Len(MyFiles) + 1 - InStr(1, MyFiles, "/"))
        GenBank_File_Path = MyFiles
        
        If GenBank_File_Path <> "" Then
            Call Print_Log(0, "Annotate_sgRNAs_Browse", "Selected file: " & GenBank_File_Path, "Good")
        Else
            Call Print_Log(0, "Annotate_sgRNAs_Browse", "No file was selected!", "Neutral")
            Exit Sub
        End If
        
        If Annotator(GenBank_File_Path) = True Then
            MsgBox "Annotation is complete!", vbInformation, Tool_Name
        Else
            MsgBox "Annotation failed! Please check the Log worksheet for details.", vbInformation, Tool_Name
        End If
        
    Else
        MsgBox "Please import or provide a list of sequences to annotate!", vbInformation, Tool_Name
    End If
    
Error_Handler:
    If Err.Number <> 0 Then
        User_Notification = Err.Description
        On Error Resume Next
        Call Print_Log(0, "Annotate_sgRNAs_Browse", User_Notification, "Bad")
        MsgBox "Annotation failed! Please check the Log worksheet for details.", vbInformation, Tool_Name
    Else
        Call Print_Log(0, "Annotate_sgRNAs_Browse", "Done!", "Good")
    End If
    
    Call Defaulter
    
End Sub

Sub Annotate_sgRNAs_Download()
    
    On Error GoTo Error_Handler
    
    Call HouseKeeping
    Call Clean_Log
    'Call Check_Version
    Call Print_Log(0, "Annotate_sgRNAs_Download", "Annotate_sgRNAs_Download Started!", "Good")
    
    If Not Len(Sheets("Main").Range("Targeted_Gene")) > 1 Then
        MsgBox "Please provide a gene name!", vbInformation, Tool_Name
        Exit Sub
    End If
    
    Dim GeneID_Lib_Path As String
    Dim Line_input As String, Gene_Name As String, Coordinate_array() As String
    Dim Chromosome As String, Position_Start As Double, Position_End As Double, Chr_Strand As String
    Dim Gene_Length As Double, Species As String
    Dim Identified_Gene As Boolean, Internal_Error As Boolean
    Identified_Gene = False
    Internal_Error = False
    
    Species = UCase(Sheets("Main").Range("Species"))
    Gene_Name = UCase(Sheets("Main").Range("Targeted_Gene") + ";")
    
    If InStr(1, Species, "HUMAN") > 0 Or InStr(1, Species, "HOMO") > 0 Then
        GeneID_Lib_Path = ActiveWorkbook.Path & "/Library/" & "GeneID_Human.txt"
    ElseIf InStr(1, Species, "MOUSE") > 0 Or InStr(1, Species, "MUS") > 0 Then
        GeneID_Lib_Path = ActiveWorkbook.Path & "/Library/" & "GeneID_Mouse.txt"
    Else
        Internal_Error = True
        User_Notification = "Couldn't recognize the Species. Please check the spelling (acceptable input: Human/Mouse)."
        GoTo Error_Handler
    End If
    
    Temp_Text = Dir(GeneID_Lib_Path)
    
    If Temp_Text = "" Then
        User_Notification = "Couldn't locate GeneID_Lib_Path: " & GeneID_Lib_Path
        Call Print_Log(0, "Annotate_sgRNAs_Download", User_Notification, "Bad")
    Else
        User_Notification = "GeneID_Lib_Path: " & GeneID_Lib_Path
        Call Print_Log(0, "Annotate_sgRNAs_Download", User_Notification, "Good")
    End If
    
    Open GeneID_Lib_Path For Input As #3
    
    Do Until EOF(3)
Next_Loop:
        Line Input #3, Line_input
            If Left(UCase(Line_input), InStr(1, Line_input, ";")) = Gene_Name Then
                Coordinate_array() = Split(Line_input, ";")
                Chromosome = Coordinate_array(3)
                Position_Start = Val(Coordinate_array(4))
                Position_End = Val(Coordinate_array(5))
                Chr_Strand = Coordinate_array(6)
                Gene_Length = Position_End - Position_Start
                Identified_Gene = True
                Exit Do
            End If
    Loop
    Close #3
    
    If Identified_Gene = False Then
        Internal_Error = True
        User_Notification = "Couldn't recognize the Gene Symbol. Please check the spelling and make sure to use an official symbol."
        GoTo Error_Handler
    End If
    
    If (Position_End - Position_Start) > 300000 Then
        User_Notification = "Gene length is above 300K! Download and annotations will be only for the first 300K of the gene!"
        Call Print_Log(0, "Annotate_sgRNAs_Download", User_Notification, "Note")
        MsgBox User_Notification, vbInformation, Tool_Name
        Position_End = Position_Start + 300000
    ElseIf (Position_End - Position_Start) < 1 Then
        User_Notification = "Gene length: " & Str(Position_End - Position_Start)
        Call Print_Log(0, "Annotate_sgRNAs_Download", User_Notification, "Bad")
    Else
        User_Notification = "Gene length: " & Str(Position_End - Position_Start)
        Call Print_Log(0, "Annotate_sgRNAs_Download", User_Notification, "Good")
    End If
    
    'Create the link and download the file.
    Dim GenBank_URL As String, Promoter_Len As Long
    Promoter_Len = Abs(Sheets("Main").Range("Promoter_Length"))
    
    If Promoter_Len > 200000 Then
        User_Notification = "The provided promoter length is " & Str(Promoter_Len)
        Call Print_Log(0, "Annotate_sgRNAs_Download", User_Notification, "Bad")
        
        Promoter_Len = 200000
        Sheets("Main").Range("Promoter_Length") = Promoter_Len
        User_Notification = "The maximum allowed promoter length is 200K. Promoter length adjusted accordingly."
        Call Print_Log(0, "Annotate_sgRNAs_Download", User_Notification, "Bad")
    End If

    If Chr_Strand = "plus" Then
        Position_Start = Position_Start - Promoter_Len
        GenBank_URL = "https://www.ncbi.nlm.nih.gov/sviewer/viewer.cgi?tool=portal&save=file&log$=seqview&db=nuccore&report=genbank&id=" & Chromosome & "&from=" & Position_Start & "&to=" & Position_End & "&"
    ElseIf Chr_Strand = "minus" Then
        Position_End = Position_End + Promoter_Len
        GenBank_URL = "https://www.ncbi.nlm.nih.gov/sviewer/viewer.cgi?tool=portal&save=file&log$=seqview&db=nuccore&report=genbank&id=" & Chromosome & "&from=" & Position_Start & "&to=" & Position_End & "&strand=on&conwithfeat=on&basic_feat=on&withparts=on"
    End If
    
    GenBank_URL = Replace(GenBank_URL, " ", "")
    
    Call Print_Log(0, "Annotate_sgRNAs_Download", GenBank_URL, "Good")
    
    ThisWorkbook.FollowHyperlink GenBank_URL
    
'    Dim Tempo_File_Address As String
'
'    Tempo_File_Address = ActiveWorkbook.Path & "/Library/" & Gene_Name & ".gb"
'
'    If Download_File(GenBank_URL, Tempo_File_Address) = True Then
'
'            If Count_Records = True Then
'
'                If Annotator(Tempo_File_Address) = True Then
'                    MsgBox "Annotation is complete!", vbInformation, Tool_Name
'                Else
'                    MsgBox "Annotation failed! Please check the Log worksheet for details.", vbInformation, Tool_Name
'                End If
'
'            Else
'
'                Name Tempo_File_Address As (ActiveWorkbook.Path & "/" & Range("Targeted_Gene") & ".gb")
'                Kill (Tempo_File_Address)
'                User_Notification = "Downloaded file: " & ActiveWorkbook.Path & "/" & Gene_Name & ".gb"
'                Call Print_Log(0, "Annotate_sgRNAs_Download", User_Notification, "Good")
'                MsgBox "Nothing annotated on the genbank file!", vbInformation, Tool_Name
'                Exit Sub
'
'            End If
'
'    Else
'        MsgBox "Downloading the genbank file failed! Please check the Log worksheet.", vbInformation, Tool_Name
'    End If
'
'    Kill (Tempo_File_Address)

Error_Handler:
    On Error Resume Next
    If Err.Number <> 0 Then
        User_Notification = Err.Description
        Call Print_Log(0, "Annotate_sgRNAs_Download", User_Notification, "Bad")
        MsgBox "Downloading the genbank file failed! Please check the Log worksheet.", vbInformation, Tool_Name
    ElseIf Internal_Error = True Then
        Call Print_Log(0, "Annotate_sgRNAs_Download", User_Notification, "Bad")
        MsgBox User_Notification, vbInformation, Tool_Name
    Else
        Call Print_Log(0, "Annotate_sgRNAs_Download", "Done!", "Good")
    End If
    
    Call Defaulter

End Sub

