Attribute VB_Name = "mLogTest"
Option Explicit
Option Compare Binary
' ----------------------------------------------------------------------
' Standard Module mLogTest: Individual tests plus a Regression-Test
' ========================= which combines them all.
'
' Note: Because the clsTestAid class module itself uses the clsLog class
'       module it cannot be used to support testing of this component!
'
' W. Rauschenberger, Berlin Jun 2024
' ----------------------------------------------------------------------
Private bRegTestFailed                  As Boolean
Private sRegTestResult                  As String
Private FSo                             As New FileSystemObject
Private lLineExpected                   As Long
Private lLineResult                     As Long
Private sExpected                       As String
Private sResultExpected_FileFullName    As String
Private sResult_FileFullName            As String
Private sResult                         As String
Private sLineExpected                   As String
Private sLineResult                     As String
Private TestAid                         As New clsTestAid
Private Log                             As clsLog

#If Not mMsg = 1 Then
    ' -------------------------------------------------------------------------------
    ' The 'minimum error handling' aproach implemented with this module and
    ' provided by the ErrMsg function uses the VBA.MsgBox to display an error
    ' message which includes a debugging option to resume the error line
    ' provided the Conditional Compile Argument 'Debugging = 1'.
    ' This declaration allows the mTrc module to work completely autonomous.
    ' It becomes obsolete when the mMsg/fMsg module is installed which must
    ' be indicated by the Conditional Compile Argument MsgComp = 1.
    ' See https://github.com/warbe-maker/Common-VBA-Message-Service
    ' -------------------------------------------------------------------------------
    Private Const vbResumeOk As Long = 7 ' Buttons value in mMsg.ErrMsg (pass on not supported)
    Private Const vbResume   As Long = 6 ' return value (equates to vbYes)
#End If

Private Property Get SplitStr(ByRef s As String)
' ------------------------------------------------------------------------------
' Returns the split string in string (s) used by VBA.Split() to turn the string
' into an array.
' ------------------------------------------------------------------------------
    If InStr(s, vbCrLf) <> 0 Then SplitStr = vbCrLf _
    Else If InStr(s, vbLf) <> 0 Then SplitStr = vbLf _
    Else If InStr(s, vbCr) <> 0 Then SplitStr = vbCr
End Property

Private Function AppErr(ByVal app_err_no As Long) As Long
' ------------------------------------------------------------------------------
' Ensures that a programmed (i.e. an application) error number never conflicts
' with VB runtime error. Thr function returns a given positive number
' (app_err_no) with the vbObjectError added - which turns it to negative. When
' the provided number is negative it returns the original positive "application"
' error number e.g. for being used with an error message.
' ------------------------------------------------------------------------------
    If app_err_no >= 0 Then AppErr = app_err_no + vbObjectError Else AppErr = Abs(app_err_no - vbObjectError)
End Function

Private Function ArrayIsAllocated(a_v As Variant) As Boolean
    On Error Resume Next
    ArrayIsAllocated = Not IsError(UBound(a_v))
End Function

Private Function ArrayTrimmed(ByVal t_v As Variant) As Variant
' ----------------------------------------------------------------------------
'
' ----------------------------------------------------------------------------
    Dim i   As Long
    Dim k   As Long
    Dim arr As Variant
    
    If Not ArrayIsAllocated(t_v) Then
        Exit Function
    End If
    
    '~~ Get first code line not empty
    For i = LBound(t_v) To UBound(t_v)
        If Len(Trim(t_v(i))) > 0 Then Exit For
    Next i
    
    If i < 0 Then
    Else
        '~~ Mode all items up
        For i = i To UBound(t_v)
            t_v(k) = t_v(i)
            k = k + 1
        Next i
        arr = t_v
        '~~ Eliminate trailing empty items
        Do While Trim(arr(UBound(arr))) = vbNullString And UBound(arr) > 0
            If UBound(arr) > 0 _
            Then ReDim Preserve arr(UBound(arr) - 1) _
            Else Exit Do
        Loop
        If Not Trim(arr(UBound(arr))) = vbNullString Then
            ArrayTrimmed = arr
        Else
        End If
    End If
    
End Function

Private Sub BoP(ByVal b_proc As String, ParamArray b_arguments() As Variant)
' ------------------------------------------------------------------------------
' Common 'Begin of Procedure' interface for the 'Common VBA Error Services' and
' the 'Common VBA Execution Trace Service' (only in case the first one is not
' installed/activated). The services, when installed, are activated by the
' | Cond. Comp. Arg.             | Installed component |
' |------------------------------|---------------------|
' | ExecTraceBymTrc = 1          | mTrc                |
' | ExecTraceByclsTrc = 1        | clsTrc              |
' | ErHComp = 1                  | mErH                |
' I.e. both components are independant from each other!
' Note: This procedure is obligatory for any VB-Component using either the
'       the 'Common VBA Error Services' and/or the 'Common VBA Execution Trace
'       Service'.
' ------------------------------------------------------------------------------
    Dim s As String
    If UBound(b_arguments) >= 0 Then s = Join(b_arguments, ",")

#If mErH = 1 Then
    '~~ The error handling also hands over to the mTrc/clsTrc component when
    '~~ either of the two is installed.
    mErH.BoP b_proc, s
#ElseIf clsTrc = 1 Then
    '~~ mErH is not installed but the mTrc is
    Trc.BoP b_proc, s
#ElseIf mTrc = 1 Then
    '~~ mErH neither mTrc is installed but clsTrc is
    mTrc.BoP b_proc, s
#End If

End Sub

Private Sub EoP(ByVal e_proc As String, Optional ByVal e_inf As String = vbNullString)
' ------------------------------------------------------------------------------
' Common 'End of Procedure' interface for the 'Common VBA Error Services' and
' the 'Common VBA Execution Trace Service' (only in case the first one is not
' installed/activated).
' Note 1: The services, when installed, are activated by the
'         | Cond. Comp. Arg.             | Installed component |
'         |------------------------------|---------------------|
'         | ExecTraceBymTrc = 1          | mTrc                |
'         | ExecTraceByclsTrc = 1        | clsTrc              |
'         | ErHComp = 1                  | mErH                |
'         I.e. both components are independant from each other!
' Note 2: This procedure is obligatory for any VB-Component using either the
'         the 'Common VBA Error Services' and/or the 'Common VBA Execution
'         Trace Service'.
' ------------------------------------------------------------------------------
#If mErH = 1 Then
    '~~ The error handling also hands over to the mTrc component when 'ExecTrace = 1'
    '~~ so the Else is only for the case the mTrc is installed but the merH is not.
    mErH.EoP e_proc
#ElseIf clsTrc = 1 Then
    Trc.EoP e_proc, e_inf
#ElseIf mTrc = 1 Then
    mTrc.EoP e_proc, e_inf
#End If

End Sub

Private Function ErrMsg(ByVal err_source As String, _
               Optional ByVal err_no As Long = 0, _
               Optional ByVal err_dscrptn As String = vbNullString, _
               Optional ByVal err_line As Long = 0) As Variant
' ------------------------------------------------------------------------------
' Universal error message display service which displays:
' - a debugging option button
' - an "About:" section when the err_dscrptn has an additional string
'   concatenated by two vertical bars (||)
' - the error message either by means of the Common VBA Message Service
'   (fMsg/mMsg) when installed (indicated by Cond. Comp. Arg. `mMsg = 1` or by
'   means of the VBA.MsgBox in case not.
'
' Uses: AppErr  For programmed application errors (Err.Raise AppErr(n), ....)
'               to turn them into a negative and in the error message back into
'               its origin positive number.
'
' W. Rauschenberger Berlin, Jan 2024
' See: https://github.com/warbe-maker/VBA-Error
' ------------------------------------------------------------------------------
#If mErH = 1 Then
    '~~ When Common VBA Error Services (mErH) is availabel in the VB-Project
    '~~ (which includes the mMsg component) the mErh.ErrMsg service is invoked.
    ErrMsg = mErH.ErrMsg(err_source, err_no, err_dscrptn, err_line): GoTo xt
    GoTo xt
#ElseIf mMsg = 1 Then
    '~~ When (only) the Common Message Service (mMsg, fMsg) is available in the
    '~~ VB-Project, mMsg.ErrMsg is invoked for the display of the error message.
    ErrMsg = mMsg.ErrMsg(err_source, err_no, err_dscrptn, err_line): GoTo xt
    GoTo xt
#End If
    '~~ When neither of the Common Component is available in the VB-Project
    '~~ the error message is displayed by means of the VBA.MsgBox
    Dim ErrBttns    As Variant
    Dim ErrAtLine   As String
    Dim ErrDesc     As String
    Dim ErrLine     As Long
    Dim ErrNo       As Long
    Dim ErrSrc      As String
    Dim ErrText     As String
    Dim ErrTitle    As String
    Dim ErrType     As String
    Dim ErrAbout    As String
        
    '~~ Obtain error information from the Err object for any argument not provided
    If err_no = 0 Then err_no = Err.Number
    If err_line = 0 Then ErrLine = Erl
    If err_source = vbNullString Then err_source = Err.Source
    If err_dscrptn = vbNullString Then err_dscrptn = Err.Description
    If err_dscrptn = vbNullString Then err_dscrptn = "--- No error description available ---"
    
    '~~ Consider extra information is provided with the error description
    If InStr(err_dscrptn, "||") <> 0 Then
        ErrDesc = Split(err_dscrptn, "||")(0)
        ErrAbout = Split(err_dscrptn, "||")(1)
    Else
        ErrDesc = err_dscrptn
    End If
    
    '~~ Determine the type of error
    Select Case err_no
        Case Is < 0
            ErrNo = AppErr(err_no)
            ErrType = "Application Error "
        Case Else
            ErrNo = err_no
            If err_dscrptn Like "*DAO*" _
            Or err_dscrptn Like "*ODBC*" _
            Or err_dscrptn Like "*Oracle*" _
            Then ErrType = "Database Error " _
            Else ErrType = "VB Runtime Error "
    End Select
    
    If err_source <> vbNullString Then ErrSrc = " in: """ & err_source & """"   ' assemble ErrSrc from available information"
    If err_line <> 0 Then ErrAtLine = " at line " & err_line                    ' assemble ErrAtLine from available information
    ErrTitle = Replace(ErrType & ErrNo & ErrSrc & ErrAtLine, "  ", " ")         ' assemble ErrTitle from available information
       
    ErrText = "Error: " & vbLf & ErrDesc & vbLf & vbLf & "Source: " & vbLf & err_source & ErrAtLine
    If ErrAbout <> vbNullString Then ErrText = ErrText & vbLf & vbLf & "About: " & vbLf & ErrAbout
    ErrBttns = vbYesNo
    ErrText = ErrText & vbLf & vbLf & "Debugging:" & vbLf & "Yes    = Resume Error Line" & vbLf & "No     = Terminate"
    ErrMsg = MsgBox(Title:=ErrTitle, Prompt:=ErrText, Buttons:=ErrBttns)
xt:
End Function

Private Function ErrSrc(ByVal sProc As String) As String
    ErrSrc = "mLogTest" & "." & sProc
End Function

Private Function FileAsArray(ByVal f_file_full_name As String) As Variant
' ----------------------------------------------------------------------------
' Returns a file's (f_file) records/lines as array.
' Note when copied: Originates in mVarTrans
'                   See https://github.com/warbe-maker/Excel_VBA_VarTrans
' ----------------------------------------------------------------------------
    Const PROC = "FileAsArray"
    
    On Error GoTo eh
    FileAsArray = Split(FileAsString(f_file_full_name), vbCrLf)
    
xt: Exit Function

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Private Function FileAsString(ByVal f_file_full_name As String) As String

    Open f_file_full_name For Input As #1
    FileAsString = Input$(lOf(1), 1)
    Close #1

End Function

Private Function Min(ParamArray va() As Variant) As Variant
' --------------------------------------------------------
' Returns the minimum (smallest) of all provided values.
' --------------------------------------------------------
    Dim v As Variant
    
    Min = va(LBound(va)): If LBound(va) = UBound(va) Then Exit Function
    For Each v In va
        If v < Min Then Min = v
    Next v
    
End Function

Private Sub Prepare()

    If Not TestAid.ModeRegression Then
        Set TestAid = Nothing
        Set TestAid = New clsTestAid
    End If
    TestAid.TestedComp = "clsLog"
    TestAid.TestFileExtention = "log"
    Set Log = Nothing
    Set Log = New clsLog

End Sub

Private Sub ProvideNewTraceLogFile(ByVal p_name As String)
    Dim s As String
    
    s = TestAid.TestFolder & "\" & p_name
    With FSo
        If .FileExists(s) Then .DeleteFile s
    End With
    mTrc.FileFullName = s
    
End Sub

Private Function StringTrimmed(ByVal s_s As String, _
                      Optional ByRef s_as_dict As Dictionary = Nothing) As String
' ----------------------------------------------------------------------------
' Returns the code (s_s) provided as a single string without leading and
' trailing empty lines. When a Dictionary is provided the string is returned
' as items with the line number as key.
' ----------------------------------------------------------------------------
    Dim s As String
    Dim i As Long
    Dim v As Variant
    
    s = s_s
    '~~ Eliminate leading empty code lines
    Do While Left(s, 2) = vbCrLf
        s = Right(s, Len(s) - 2)
    Loop
    '~~ Eliminate trailing eof
    If Right(s, 1) = VBA.Chr(26) _
    Then s = Left(s, Len(s) - 1)
    
    '~~ Eliminate trailing empty code lines
    Do While Right(s, 2) = vbCrLf
        s = Left(s, Len(s) - 2)
    Loop
    
    StringTrimmed = s
    If Not s_as_dict Is Nothing Then
        With s_as_dict
            For Each v In Split(s, vbCrLf)
                i = i + 1
                .Add i, v
            Next v
        End With
    End If
    
End Function

Private Sub Test_000_Regression()
' ------------------------------------------------------------------------------
' Please note: This test includes the result assertion which is the result from
'              Test_100_All - when ok - saved to the file
'              Test_000_RegressionResultExpected.log in this projects parent folder.
' ------------------------------------------------------------------------------
    Const PROC = "Test_000_Regression"
    
    On Error GoTo eh
    Set TestAid = Nothing
    Set TestAid = New clsTestAid
    
    With TestAid
        .ModeRegression = True
        mErH.Regression = True
        ProvideNewTraceLogFile "Test-Regression.trc"
        
        BoP ErrSrc(PROC)
        Test_010_Align_Normal
        Test_020_Align_Col_Arranged
        Test_100_All
        
        EoP ErrSrc(PROC)
        mErH.Regression = False
        
        .ResultSummaryLog
        .CleanUp ' cleanup obsolete files

    End With
    
xt: Exit Sub

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub Test_001_Guide()

    Dim s1 As String
    Dim s2 As String
    Dim i  As Long
    
    If Not TestAid Is Nothing Then Set TestAid = Nothing
    Set TestAid = New clsTestAid
    
    With TestAid
        .TestId = "GuideTest-1"
        .TestedComp = "clsTestAid"
        .TestedProc = "Guide, GuideDone"
        .Verification = "Basic functionality"
        
        .Guide "This test guidance message should appear at tthe top left position." & vbLf & _
               "Move the message to any other porition and continue!"
        Stop
        
        .Guide "This message should appear at the position the first one had been left."
        Stop
        
        s1 = "Lorem ipsum sic transit ... "
        For i = 1 To 50
            s2 = s2 & s1
        Next i
        .Guide s2
        Stop
                
        .GuideTerminate
        
        .Guide "This message should have poped up at the last position it had been displayed." & vbLf & vbLf & s2
        Stop
        
    End With
    
    Set TestAid = Nothing
    
End Sub

Private Sub Test_010_Align_Normal()

    Dim sResult As String
    Dim lWidth  As Long
    
    Prepare
    With TestAid
        .TestHeadLine = "Alignment straight (not column arranged)"
        .TestedProc = "Align"
        .TestedType = "Function"
        
        ' =====================================================================
        .TestId = "010-1"
        .Verification = "Align left, margin= "" "" (ignored!), fill="" "" (final width = 10)"
        .Result = ">" & Log.Align("xx", "L", 10, " ", " ", False) & "<"
        .ResultExpected = ">xx        <"
        
        ' =====================================================================
        .TestId = "010-2"
        .Verification = "Align left, margin= "" "" (ignored!), fill="" "" (final width = 10)"
        .Result = ">" & Log.Align("xxxxxxxxxxxx", "L", 10, " ", " ", False) & "<"
        .ResultExpected = ">xxxxxxxxxx<"
        
        ' =====================================================================
        .TestId = "010-3"
        .Verification = "Align right, margin="" "" (ignored), fill="" "" (final length = 10)"
        .Result = ">" & Log.Align("x", "R", 10, " ", " ", False) & "<"
        .ResultExpected = ">         x<"
        
        ' =====================================================================
        .TestId = "010-4"
        .Verification = "Align centered, margin="""", fill=""-"" (final width is 10)"
        .Result = ">" & Log.Align("x", "C", 10, "-", "", False) & "<"
        .ResultExpected = ">----x-----<" ' final width is 12 (10 plus one - fill at the left and one at the right)
              
        ' =====================================================================
        .TestId = "010-5"
        .Verification = "Align centered, margin="""", fill="" -"" (final width is 10)"
        .Result = ">" & Log.Align("x", "C", 10, " -", "", False) & "<"
        .ResultExpected = ">--- x ----<" ' final width is 12 (10 plus one - fill at the left and one at the right)
        ' =====================================================================
        
        .TestId = "010-6"
        .Verification = "Align left, margin="""", fill="".:"" (final width is 10)"
        .Result = ">" & Log.Align("xx", "L", 10, ".:", , False) & "<"
        .ResultExpected = ">xx.......:<" ' final width is 12 (10 plus - at least - one fill at the right)
            
        ' =====================================================================
        .TestId = "010-7"
        .Verification = "Align left, margin="""", fill=""."" (final width is 10"
        .Result = ">" & Log.Align("xxx", "L", 10, ".", " ", False) & "<"
        .ResultExpected = ">xxx.......<" ' final width is 12 (10 plus one . fill at the right)
    
        .CleanUp "*-Result-1.log", "*-Result-2.log"
    End With
    
End Sub

Private Sub Test_020_Align_Col_Arranged()

    Dim sResult As String
    Dim lWidth  As Long
    
    Prepare
    With TestAid
        .TestHeadLine = "Alignments column arranged"
        .TestedProc = "Align"
        .TestedType = "Function"
        
        ' =====================================================================
        .TestId = "020-1"
        .Verification = "Align left, cols arragend, margin= "" "", fill="" "" (final width = 12)"
        .Result = "|" & Log.Align("xx", "L", 10, " ", " ", True) & "|"
        .ResultExpected = "| xx         |"
        
        ' =====================================================================
        .TestId = "020-2"
        .Verification = "Align right, cols arranged, margin="" "", fill="" "" (final length = 12)"
        .Result = ">" & Log.Align("x", "R", 10, " ", " ", True) & "<"
        .ResultExpected = ">          x <"
        
        ' =====================================================================
        .TestId = "020-3"
        .Verification = "Align centered, cols arranged, margin="" "", fill=""-"" (final width is 14)"
        .Result = "|" & Log.Align("x", "C", 10, "-", " ", True) & "|"
        .ResultExpected = "| -----x------ |" ' final width is 12 (10 plus one - fill at the left and one at the right)
        
        ' =====================================================================
        .TestId = "020-4"
        .Verification = "Align centered, cols arranged, margin="" "", fill="" -"" (final width is 14)"
        .Result = "|" & Log.Align("x", "C", 10, " -", " ", True) & "|"
        .ResultExpected = "| ----- x ------ |" ' final width is 12 (10 plus one - fill at the left and one at the right)
        
        ' =====================================================================
        .TestId = "020-5"
        .Verification = "Align left, cols arranged, fill="".:"", colls arranged (final width is 12)"
        .Result = Log.Align("xx", "L", 10, ".:", " ", True)
        .ResultExpected = " xx.........: " ' final width is 12 (10 plus - at least - one fill at the right)
            
        ' =====================================================================
        .TestId = "020-5"
        .Verification = "Align left, filled with ""."", considered cols arranged, final width is 11"
        .Result = Log.Align("xxx", "L", 10, ".", " ", True)
        .ResultExpected = " xxx........ " ' final width is 12 (10 plus one . fill at the right)
    
        .CleanUp "*-Result-1.log", "*-Result-2.log"
    End With
    
End Sub

Private Sub Test_100_All()
' ------------------------------------------------------------------------------
' Regression-Test: The description of the individual test is either provided
'                  by a Title for a series of log-entries or - when no Title
'                  is specified - by the log-entries themselve.
' ------------------------------------------------------------------------------
    Const PROC = "Test_100_All"
    
    On Error GoTo eh
    Dim Log             As New clsLog
    Dim bTimeStamp      As Boolean: bTimeStamp = True
    Dim lLines          As Long
    Dim bAsserted       As Boolean
    
    If Not mErH.Regression Then ProvideNewTraceLogFile "Test.trc"
    BoP ErrSrc(PROC)
            
    Prepare
    With TestAid
        .ExcludeFromComparison = "??-??-??-??:??:?? " ' like string excluded from result/result expected string comparison
        .TestedProc = "Title, Entry"
        .TestHeadLine = "All log services"
        
        ' ====================================================
        .TestId = "100-1"
        .Verification = "1. A log entry with 4 line title, 2. One log entry as new log"
        With Log
            .FileFullName = TestAid.NameTestResultFile
            .NewFile
            .WithTimeStamp bTimeStamp
            .KeepLogs = 11 ' for this test in order not to obstruct the result output
            
            .Title "L", "01-1 Title test: " _
                 , "- 4 title lines as ParamAray (lines are comma delimited string)" _
                 , "- 2 Single line log entries" _
                 , "- Title left adjusted by means of a trailing space with the first title line"
            .Entry "01-1 1. Single string, new log, Single string, new log."
            .NewLog
            .Entry "01-1 2. Single string, new log, no title. "
        End With
        .Result = .TestResultFile
'        .DsplyFile .NameTestResultFile
        .ResultExpected = .TestResultExpectedFile
        
        ' ====================================================
        .TestId = "100-2"
        .Verification = "New log file, title centered filled with -, 2 log entires"
        With Log
            .FileFullName = TestAid.NameTestResultFile
            .NewFile
            .WithTimeStamp bTimeStamp
            .KeepLogs = 11 ' for this test in order not to obstruct the result output
            .Title "C -"
            .Title "01-2 Title test:"
            .Title "(aligned centered filled with -)"
            .Entry "01-2 1. Single string, new log, Single string, new log."
            .Entry "01-2 2. Single string without any width limit"
        End With
        .Result = .TestResultFile
'        .DsplyFile .NameTestResultFile
        .ResultExpected = .TestResultExpectedFile
            
        ' ====================================================
        .TestId = "100-3"
        .Verification = "Title width adjusted, indicates new log"
        '~~ To continue with result of previous result the
        '~~ result file from the previous test is copied for this test
        FSo.CopyFile .TestFolder & "\Test-100-2-Result-1.log", TestAid.NameTestResultFile
        With Log
            .FileFullName = TestAid.NameTestResultFile
            .Title "C -", "01-3 Title: Regression test case 01-3"
            .Title "(centered, filled with leading and trailing - )"
            .Entry "01-3 1. Single string, new log. This is an extra long text to force all title lines with fill characters"
            .Entry "01-3 2. Single string without any width limit"
        End With
        .Result = .TestResultFile
'        .DsplyFile .NameTestResultFile
        .ResultExpected = .TestResultExpectedFile
                            
        ' ====================================================
        .TestId = "100-4"
        .Verification = "New log file, items arranged in columns, no column header"
        With Log
            .FileFullName = TestAid.NameTestResultFile
            .NewFile
            .ColsSpecs = "R4, C6, L10, L20" ' essentail to indicate the number of columns when no headers do it implicitly
            .Title "Test : Items arranged in columns (no header)"
            .ColsItems "02-1", "xxxx", "yyyyyy", "No Title! Alignments: R, C, R, L; Rightmost column without width limit"
            .ColsItems "02-1", "xxxx", "yyyyy", "Note that the rightmost column, when aligned Left (the default) is not truncated!"
        End With
        .Result = .TestResultFile
'        .DsplyFile .NameTestResultFile
        .ResultExpected = .TestResultExpectedFile
         
        ' ====================================================
        .TestId = "100-5"
        .Verification = "Items arranged in columns (with header)"
        With Log
            .FileFullName = TestAid.NameTestResultFile
            .NewFile
            .Title "C -", "Test 4: Items arranged in columns (with header) "
            .ColsSpecs = "R2, L10, R10, L20"
            .ColsDelimiter = "|"
            '~~ The below first header line implicitely specifies the column alignments
            '~~ 1: Right, 2: Center, 3: Center, 4: Left
            '~~ an this specs also apply for the Entry items.
            .ColsHeader = "Nr, Item, Item, Item"
            .ColsHeader = "  , 1   , 2   , 4 (no width limit)"
            '~~ Note: The above first header line implicitly specifies the columns alignment
            '~~       Only in case no headers had been specified prior the first Entry
            '~~       the first Entry line implicitely specifies the columns alignment
            '~~       provided none had already been specified explicitley
            .ColsItems "3", "xxxxx", "yyyyyy", "Rightmost column without width limit! (this first line implicitly indicates the columns width for being considered by the header) "
            .ColsItems "3", "xxx", "yyyy", "         zzzzzz (note that leading spaces are preserved when/because the first line implicitly indicated left adjusted)"
            .ColsItems "3", "x", "yyyyy", "zzzzzz"
        End With
        .Result = .TestResultFile
'        .DsplyFile .NameTestResultFile
        .ResultExpected = .TestResultExpectedFile
         
        ' ====================================================
        .TestId = "100-6"
        .Verification = "Items arranged in columns (fill "" .:"", no header)"
        With Log
            .FileFullName = TestAid.NameTestResultFile
            .NewFile
            .Title "C =", "Test 4: Items arranged in columns (with header) "
            .ColsSpecs = "R2, L50 .:, L10, L20"
            .ColsItems "3", "xxxxx", "yyyyyy", "Rightmost column no width limit!"
            .ColsItems "3", "xxx", "yyyy", "  - Header and title is adjusted with first (above) line"
            .ColsItems "3", "x", "yyyyy", "  - Leading spaces are preserved with left alignment, rigth are with right alignment)"
        End With
        .Result = .TestResultFile
'        .DsplyFile .NameTestResultFile
        .ResultExpected = .TestResultExpectedFile
    
        ' ====================================================
        .TestId = "100-7"
        .Verification = "Items arranged in columns (fill "".:"", cols with empty items)"
        With Log
            .FileFullName = TestAid.NameTestResultFile
            .NewFile
            .Title "C -", "Test 4: Items arranged in columns (with header) "
            .ColsHeader = " No., Item, Comment"
            .ColsSpecs = "R4, L50 .:, L10"
            .ColsDelimiter = " "
            .ColsItems "1", "xxxxx", "Rightmost column without width limit! (this first line implicitly indicates the columns width for being considered by the header) "
            .ColsItems "2", "xxx", "  - zzzzzz (note that leading spaces are preserved with left alignment - as rigth are with right alignment)"
            .ColsItems "3", "x", "zzzzzz"
            .ColsItems "", "", "Additional comment to the above item"
        End With
        .Result = .TestResultFile
'        .DsplyFile .NameTestResultFile
        .ResultExpected = .TestResultExpectedFile
        
        .CleanUp "*-Result-1.log", "*-Result-2.log"
    End With

xt: EoP ErrSrc(PROC)
    Exit Sub
    
eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

