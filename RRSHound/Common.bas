Attribute VB_Name = "Common"
Option Explicit

' Note - Use error handlers in procedures that call these functions

Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
          (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)

Public Declare Function AllocStrB Lib "oleaut32" Alias "SysAllocStringByteLen" _
          (ByVal lpszStr As Long, ByVal lLenB As Long) As Long

Public Declare Function OSWinHelp Lib "user32" Alias "WinHelpA" _
          (ByVal hWnd As Long, ByVal lpHelpFile As String, _
           ByVal wCommand As Long, ByVal dwData As Long) As Long

Public Const ZERO As Long = 0
Public Const ONE As Long = 1
Public Const TWO As Long = 2
Public Const THREE As Long = 3
Public Const FOUR As Long = 4
Public Const FIVE As Long = 5

Private Const INVALID_ARG As Long = FIVE

Public Enum HelpTypeFlags
    HelpContents = THREE
    HowToUseHelp = FOUR
    SearchForHelpOn = 261
End Enum


Private Const INVALID_FILE_ATTRIBUTES As Long = &HFFFFFFFF ' -1

Public Enum vbFileAttributes
    vbInvalidFile = -1   ' Returned INVALID_FILE_ATTRIBUTES
    vbNormal = 0         ' Normal (default for SetAttributes)
    vbReadOnly = 1       ' Read-only
    vbHidden = 2         ' Hidden
    vbSystem = 4         ' System file
    vbVolume = 8         ' Volume label
    vbDirectory = 16     ' Directory or folder
    vbArchive = 32       ' File has changed since last backup
    vbTemporary = &H100  ' 256
    vbCompressed = &H800 ' 2048
End Enum

Private Declare Function GetLongPathName Lib "kernel32" Alias "GetLongPathNameA" _
    (ByVal lpszShortPath As String, ByVal lpszLongPath As String, _
     ByVal cchBuffer As Long) As Long

Private Declare Function FormatMessage Lib "kernel32" Alias "FormatMessageA" _
    (ByVal dwFlags As Long, lpSource As Any, ByVal dwMessageId As Long, _
     ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, _
     Arguments As Long) As Long

Private Const FORMAT_MESSAGE_FROM_SYSTEM = &H1000&
Private Const sEMPTY_STR = "Empty string passed."

' Maximum allowed path length including path, filename,
' and command line arguments for NT (Intel) and Win95

Public Const DIR_SEP As String = "\"

' PI is: 3.1415926535897932384626433832795028841971694

Public Function RandomColor() As Long 'Returns a random RBG color
    Dim lRed As Long ' ©Jay
    Dim lGreen As Long
    Dim lBlue As Long
    
    Call Randomize(Timer)
    lRed = Random(ONE, 255)
    lGreen = Random(ONE, 255)
    lBlue = Random(ONE, 255)
    RandomColor = RGB(lRed, lGreen, lBlue)
End Function

' Bounced off Paul Catons InIDE
Public Function InTheIDE() As Boolean
    Debug.Assert True Xor DebugOnly(InTheIDE)
End Function

Private Function DebugOnly(fInIDE As Boolean) As Boolean
    fInIDE = True
End Function

Public Function GetAttrib(sFileSpec As String, ByVal Attrib As vbFileAttributes) As Boolean
    ' Returns True if the specified attribute(s) is currently set.
    If LenB(sFileSpec) = ZERO Then Err.Raise INVALID_ARG, , sEMPTY_STR
    GetAttrib = (GetAttributes(sFileSpec) And Attrib) = Attrib
End Function

Public Sub SetAttrib(sFileSpec As String, ByVal Attrib As vbFileAttributes, Optional fTurnOff As Boolean)
    ' Sets/clears the specified attribute(s) without affecting other attributes. You
    ' do not need to know the current state of an attribute to set it to on or off.
    If LenB(sFileSpec) = ZERO Then Err.Raise INVALID_ARG, , sEMPTY_STR
    Dim Attribs As Long: Attribs = GetAttributes(sFileSpec)
    If fTurnOff Then
        SetAttributes sFileSpec, Attribs And (Not Attrib)
    Else
        SetAttributes sFileSpec, Attribs Or Attrib
    End If
End Sub




Public Function DirExists(sPath As String, Optional fCreateIfNot As Boolean) As Boolean
    If LenB(sPath) = ZERO Then Err.Raise INVALID_ARG, , sEMPTY_STR
    Dim Attribs As Long: Attribs = GetAttributes(sPath)
    If (Attribs <> INVALID_FILE_ATTRIBUTES) Then
        DirExists = ((Attribs And vbDirectory) = vbDirectory)
    End If
    If (DirExists = False) Then
        If fCreateIfNot Then DirExists = CreatePath(sPath)
    End If
End Function

'-----------------------------------------------------------
' Creates the specified file sFileSpec.
' Returns: 2 if created, 1 if existed, 0 if error.
'-----------------------------------------------------------
Public Function CreateFile(sFileSpec As String) As Long
    If LenB(sFileSpec) = ZERO Then Err.Raise INVALID_ARG, , sEMPTY_STR
    Dim iFile As Integer, Idx As Integer, sFile As String

    On Error GoTo FailedCreateFile
    If FileExists(sFileSpec) Then
        CreateFile = ONE
    Else
        sFile = LongPathName(sFileSpec)
        Idx = InStrR(sFile, DIR_SEP)
        If (Idx > ZERO) And (Idx < Len(sFile)) Then
            If CreatePath(Left$(sFile, Idx)) Then
                iFile = FreeFile
                Open sFile For Output As #iFile
                Close #iFile
                CreateFile = TWO
    End If: End If: End If
FailedCreateFile:
End Function

'-----------------------------------------------------------
' Creates the specified directory sPath.
' Returns: 2 if created, 1 if existed, 0 if error.
'-----------------------------------------------------------
Public Function CreatePath(sPath As String) As Long
    If LenB(sPath) = ZERO Then Err.Raise INVALID_ARG, , sEMPTY_STR
    Dim sDir As String, sTemp As String, Idx As Integer

    On Error GoTo FailedCreatePath
    If (DirExists(sPath)) Then
        CreatePath = ONE
    Else
        ' Add trailing backslash if missing
        sDir = AddBackslash(LongPathName(sPath))

        ' Set Idx to the first backslash
        Idx = InStr(ONE, sDir, DIR_SEP)

        Do ' Loop and make each subdir of the path separately
            Idx = InStr(Idx + ONE, sDir, DIR_SEP)
            If (Idx <> ZERO) Then
                sTemp = Left$(sDir, Idx - ONE)
                ' Determine if this directory already exists
                If (DirExists(sTemp) = False) Then
                    ' We must create this directory
                    MkDir sTemp
                    CreatePath = TWO
                End If
            End If
        Loop Until Idx = ZERO
    End If
FailedCreatePath:
End Function

Public Function LongPathName(sPathName As String) As String
    If LenB(sPathName) = 0 Then Exit Function
    LongPathName = sPathName ' Default to the passed name
    On Error GoTo GetFailed
    Dim sPath As String, lResult As Long
    sPath = String$(MAX_PATH, vbNullChar)
    lResult = GetLongPathName(sPathName, sPath, MAX_PATH)
    If (lResult) Then LongPathName = TrimZ(sPath)
GetFailed:
End Function

Public Function TrimZ(sNullTerminated As String) As String
    If (LenB(sNullTerminated) = ZERO) Then Err.Raise INVALID_ARG, , sEMPTY_STR
    Dim Idx As Integer: Idx = InStr(sNullTerminated, vbNullChar)
    If (Idx <> ZERO) Then
        TrimZ = Left$(sNullTerminated, Idx - ONE)
    Else
        TrimZ = Trim$(sNullTerminated)
    End If
End Function

Public Function InStrR(sSrc As String, sTerm As String, _
                       Optional ByVal lLeftBound As Long = ONE, _
                       Optional ByVal lRightBound As Long, _
                       Optional CaseSensative As Boolean) As Long
    If (LenB(sSrc) = ZERO) Or (LenB(sTerm) = ZERO) Then Err.Raise INVALID_ARG
    Dim lPos As Long, lTerm As Long
    If lRightBound = ZERO Then lRightBound = Len(sSrc)
    lTerm = Len(sTerm)
    lRightBound = (lRightBound - lTerm) + ONE
    If CaseSensative Then
        lLeftBound = InStr(sSrc, sTerm)
        If lLeftBound = ZERO Then Exit Function
        For lPos = lRightBound To lLeftBound Step -ONE
            If (Mid$(sSrc, lPos, lTerm) = sTerm) Then
                InStrR = lPos
                Exit Function
            End If
        Next lPos
    Else
        Dim sText As String, sFind As String
        sText = LCase$(sSrc): sFind = LCase$(sTerm)
        lLeftBound = InStr(sText, sFind)
        If lLeftBound = ZERO Then Exit Function
        For lPos = lRightBound To lLeftBound Step -ONE
            If (Mid$(sText, lPos, lTerm) = sFind) Then
                InStrR = lPos
                Exit Function
            End If
        Next lPos
    End If
End Function

'-BuildStr-----------------------------------------------
'  This function can replace vb's string & concatenation.
'  The speed is exactly the same for simple appends:
'     sResult = sResult & "text"
'     sResult = BuildStr(sResult, "text")
'  But for more substrings this function is much faster
'  because vb's multiple appending is very slow:
'     sResult = sResult & "some" & "more" & "text"
'     sResult = BuildStr(sResult, "some", "more", "text")
'  Notice you can safely pass as an argument the variable
'  that the function is assigning back to (compiler safe).
'  You can also specify the delimiter character(s) to
'  insert between the appended substrings, and will work
'  correctly if an argument is omitted or passed empty:
'     sMsg = BuildStr("s1", , "s2", "s3", vbCrLf)
'     MsgBox BuildStr("", sMsg, , "s4", vbCrLf)
'--------------------------------------------------------
Public Function BuildStr(Str1 As String, Optional Str2 As String, Optional Str3 As String, Optional Str4 As String, Optional Delim As String) As String
    Dim LenWrk As Long, LenAll As Long
    Dim LenDlm As Long, CntDlm As Long
    Dim Len1 As Long, Len2 As Long
    Dim Len3 As Long, Len4 As Long
    Dim lpStr As Long
    Len1 = LenB(Str1): Len2 = LenB(Str2)
    Len3 = LenB(Str3): Len4 = LenB(Str4)
    LenDlm = LenB(Delim)
    If (LenDlm <> ZERO) Then
        CntDlm = -LenDlm
        If (Len1 <> ZERO) Then CntDlm = ZERO
        If (Len2 <> ZERO) Then CntDlm = CntDlm + LenDlm
        If (Len3 <> ZERO) Then CntDlm = CntDlm + LenDlm
        If (Len4 <> ZERO) Then CntDlm = CntDlm + LenDlm
    End If
    LenAll = Len1 + Len2 + Len3 + Len4 + CntDlm
    If (LenAll > ZERO) Then
        lpStr = AllocStrB(0&, LenAll)

        ' Preserve Unicode by passing StrPtr and byte count
        If (Len1 <> ZERO) Then
            CopyMemory ByVal lpStr, ByVal StrPtr(Str1), Len1
            LenWrk = Len1
        End If
        
        If (Len2 <> ZERO) Then
            If (LenDlm <> ZERO) Then If (LenWrk <> ZERO) Then GoSub InsDelim
            CopyMemory ByVal lpStr + LenWrk, ByVal StrPtr(Str2), Len2
            LenWrk = LenWrk + Len2
        End If
        
        If (Len3 <> ZERO) Then
            If (LenDlm <> ZERO) Then If (LenWrk <> ZERO) Then GoSub InsDelim
            CopyMemory ByVal lpStr + LenWrk, ByVal StrPtr(Str3), Len3
            LenWrk = LenWrk + Len3
        End If
        
        If (Len4 <> ZERO) Then
            If (LenDlm <> ZERO) Then If (LenWrk <> ZERO) Then GoSub InsDelim
            CopyMemory ByVal lpStr + LenWrk, ByVal StrPtr(Str4), Len4
        End If
    End If
    CopyMemory ByVal VarPtr(BuildStr), ByVal VarPtr(lpStr), 4&
    Exit Function
    
InsDelim:
    CopyMemory ByVal lpStr + LenWrk, ByVal StrPtr(Delim), LenDlm
    LenWrk = LenWrk + LenDlm
    Return
End Function

Public Function Padding(Value, ByVal lLength As Long, _
                       Optional sPadChar As String = "0", _
                       Optional fSuffixPadding As Boolean) As String
    If (lLength > Len(Value)) Then
        If fSuffixPadding Then
            Padding = CStr(Value) & String$(lLength - Len(Value), sPadChar)
        Else
            Padding = String$(lLength - Len(Value), sPadChar) & CStr(Value)
        End If
    Else
        Padding = Left$(CStr(Value), lLength)
    End If
End Function

Public Function Today() As Date
    Today = Date
End Function

Public Function Tomorrow() As Date
    Tomorrow = DateAdd("d", ONE, Date)
End Function

Public Function Yesterday() As Date
    Yesterday = DateAdd("d", -ONE, Date)
End Function

Public Function AddBackslash(sSpec As String) As String
    If (LenB(sSpec) = ZERO) Then Err.Raise INVALID_ARG, , sEMPTY_STR
    ' Add trailing backslash if missing
    If Right$(sSpec, ONE) <> DIR_SEP Then
        AddBackslash = sSpec & DIR_SEP
    Else
        AddBackslash = sSpec
    End If
End Function

Public Function RemoveBackslash(sSpec As String) As String
    If (LenB(sSpec) = ZERO) Then Err.Raise INVALID_ARG, , sEMPTY_STR
    ' Remove trailing backslash
    If Right$(sSpec, ONE) = DIR_SEP Then
        RemoveBackslash = Left$(sSpec, Len(sSpec) - ONE)
    Else
        RemoveBackslash = sSpec
    End If
End Function

Public Function AddPathSeperator(sSpec As String, Optional sSep As String) As String
    If (LenB(sSpec) = ZERO) Then Err.Raise INVALID_ARG, , sEMPTY_STR
    If (LenB(sSep) = ZERO) Then
        sSep = IIf(InStr(sSpec, "/"), "/", DIR_SEP)
    Else
        If InStr(sSpec, IIf(sSep = DIR_SEP, "/", DIR_SEP)) Then Err.Raise INVALID_ARG
    End If
    ' Add trailing seperator if missing
    If Right$(sSpec, ONE) <> sSep Then
        AddPathSeperator = sSpec & sSep
    Else
        AddPathSeperator = sSpec
    End If
End Function

Public Function RemovePathSeperator(sSpec As String) As String
    If (LenB(sSpec) = ZERO) Then Err.Raise INVALID_ARG, , sEMPTY_STR
    Dim sSep As String: sSep = IIf(InStr(sSpec, "/"), "/", DIR_SEP)
    ' Remove trailing seperator
    If Right$(sSpec, ONE) = sSep Then
        RemovePathSeperator = Left$(sSpec, Len(sSpec) - ONE)
    Else
        RemovePathSeperator = sSpec
    End If
End Function

Public Function TrimQuotes(sStr As String) As String
    If (LenB(sStr) <> 0) Then
        Dim s As String: s = Trim$(sStr)
        ' Remove double quotes if present
        If (Left$(s, 1) = Chr$(34)) Then
            TrimQuotes = Mid$(s, 2, Len(s) - 2)
        Else
            TrimQuotes = s
    End If: End If
End Function

Public Function Random(Optional ByVal min As Long = 1, Optional ByVal max As Long = 65535, Optional ByVal seed As Long = THREE) As Long
    Randomize (FIVE * seed) * Timer
    Random = CLng((max - min) * Rnd + min)
End Function

Public Function RandomInt(Optional ByVal seed As Long = THREE) As Integer
    Randomize (FIVE * seed) * Timer
    RandomInt = CInt((65535 * Rnd) - 32768)
End Function

Public Function RandomLong(Optional ByVal seed As Long = THREE) As Long
    Randomize (FIVE * seed) * Timer
    RandomLong = CLng((4294967295# * Rnd) - 2147483648#)
End Function

Public Function RandomSingle(Optional ByVal seed As Long = THREE) As Single
    Randomize (FIVE * seed) * Timer
    RandomSingle = CSng((6.85646E+37 * Rnd) - 3.42823E+37)
End Function

Public Function RandomDouble(Optional ByVal seed As Long = THREE) As Double
    RandomDouble = CDbl(RandomSingle(seed) * RandomSingle(seed))
End Function

Public Sub HourGlass(Optional fOn As Boolean = True)
    Static OrigPointer As Long ' 0 = vbDefault
    If fOn Then
        If Screen.MousePointer <> vbHourglass Then
            ' Save pointer and set hourglass
            OrigPointer = Screen.MousePointer
            Screen.MousePointer = vbHourglass
        End If
    Else
        ' Restore pointer
        Screen.MousePointer = OrigPointer
    End If
End Sub

Public Function GetFuncPointer(ByVal AddressOf_Func As Long) As Long
    GetFuncPointer = AddressOf_Func
End Function

Public Function SysRetCodeToMsg(ByVal APIReturnCode As Long) As String
    SysRetCodeToMsg = "Unknown System Error"

    If (APIReturnCode <> ZERO) Then
        Dim rc As Long, lFlag As Long
        Dim MsgBuf As String, BufLen As Long

        BufLen = 1024
        MsgBuf = String$(BufLen, vbNullChar)
        lFlag = FORMAT_MESSAGE_FROM_SYSTEM

        rc = FormatMessage(lFlag, ZERO&, APIReturnCode, ZERO&, MsgBuf, BufLen, ZERO&)
        If (rc <> ZERO) Then SysRetCodeToMsg = TrimZ(MsgBuf)
    Else
        SysRetCodeToMsg = "Returned success or failed to return"
    End If

End Function

Public Function Help(ByVal Me_hWnd As Long, ByVal HelpType As HelpTypeFlags) As Long
    Dim rc As Long
    ' If there is no helpfile for this project display a message to the user
    ' Set the HelpFile for your application in the Project Properties dialog
    If (HelpType = FOUR) Then
        rc = OSWinHelp(Me_hWnd, "Winhlp32.hlp", FOUR, ZERO)
    Else
        If Len(App.HelpFile) = ZERO Then
            MsgBox "Unable to display Help Contents.", vbInformation, App.Title
        Else
            rc = OSWinHelp(Me_hWnd, App.HelpFile, HelpType, ZERO)
        End If
    End If
End Function

'SplitStr''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' This is a smart Split function. If you get a list of paths from an
' open file dialog, or some other source, you don't need to identify
' the seperator/delimiter character before splitting - just call this
' function without specifying the split character for auto-splitting.
'
' This function automatically handles quoted/unquoted combinations.
'
' Of course, you can specify any char(s) to be used (and so removed)
' during a split.
'
' Renamed from SplitString (vb6) to SplitStr (vb5/6) and so recieves
' the array byref for vb5 compatability.
'
''
'' Usage
''
'' From Command string:
''
'    Dim aFiles() As String, Idx As Long
'
'    SplitStr Command, aFiles
'
'    For Idx = 0 To UBound(aFiles)
'        YourOpenAFile aFiles(Idx)
'    Next Idx
'
''
'' Open File Dialog:
''
'    Dim sFile As String, aFiles() As String, Idx As Long
'
'    ' File must exist, allow multi-select, use new style dialog
'    dlgCommon.flags = cdlOFNFileMustExist Or _
'                      cdlOFNAllowMultiselect Or _
'                      cdlOFNExplorer
'    On Error Resume Next
'    dlgCommon.ShowOpen
'    If (Err = cdlCancel) Then Exit Sub
'
'    sFile = dlgCommon.filename
'
'    If (sFile = vbNullString) Then
'        ' User didn't return a file name
'        Exit Sub
'    ElseIf (InStr(sFile, Chr$(0)) = 0) Then
'        ' User chose a single file name
'        YourOpenAFile sFile
'    Else
'        ' User chose multiple file names, Chr(0) is the
'        ' Nul character used to seperate the path and the
'        ' following file names
'
'        SplitStr sFile, aFiles
'
'        For Idx = 1 To UBound(aFiles)
'            YourOpenAFile aFiles(0) & "\" & aFiles(Idx)
'        Next Idx
'    End If
'
''
'' From custom sFileText string:
''
'    Dim aLines() As String, Idx As Long
'
'    SplitStr sFileText, aLines, vbCrLf
'
'    For Idx = 0 To UBound(aLines)
'        YourFileByLine aLines(Idx)
'    Next Idx
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SplitStr(sStr As String, aStrsOut() As String, Optional ByVal sSplitChr As String)
    If (LenB(Trim$(sStr)) = ZERO) Then Err.Raise INVALID_ARG, , sEMPTY_STR
    Dim iLength As Integer, idxArr As Integer
    Dim idx1 As Integer, idx2 As Integer
    Dim s As String
    s = Trim$(sStr)
    idx1 = ONE
    If sSplitChr = Chr$(34) Then sSplitChr = vbNullString ' Handle double quotes below
    If LenB(sSplitChr) <> ZERO Then
       If (InStr(s, sSplitChr)) Then ' Seperated by specified char(s)
          Do While idx1 < Len(s)
              idx2 = InStr(idx1, s, sSplitChr)
              If (idx2 = ZERO) Then idx2 = Len(s) + ONE
              iLength = idx2 - idx1
              ReDim Preserve aStrsOut(ZERO To idxArr) As String
              aStrsOut(idxArr) = TrimQuotes(Mid$(s, idx1, iLength))
              idxArr = idxArr + ONE
              idx1 = idx2 + Len(sSplitChr)
          Loop
       Else
          ReDim aStrsOut(ZERO To ZERO) As String
          aStrsOut(ZERO) = TrimQuotes(s)
       End If
    ElseIf (InStr(s, vbNullChar)) Then ' Seperated by Nulls
        Do While idx1 < Len(s)
            idx2 = InStr(idx1, s, vbNullChar)
            If (idx2 = ZERO) Then idx2 = Len(s) + ONE
            iLength = idx2 - idx1
            ReDim Preserve aStrsOut(ZERO To idxArr) As String
            aStrsOut(idxArr) = TrimQuotes(Mid$(s, idx1, iLength))
            idxArr = idxArr + ONE
            idx1 = idx2 + ONE
        Loop
    Else ' Seperated by Spaces
        Dim iSpc As Integer, iQuot As Integer
        iSpc = InStr(s, Chr$(32)) ' Space char
        iQuot = InStr(s, Chr$(34)) ' Quote char
        If (iSpc = ZERO) Then
            ReDim aStrsOut(ZERO To ZERO) As String
            aStrsOut(ZERO) = TrimQuotes(s)
        ElseIf (iQuot = ZERO) Then ' And (iSpc <> ZERO)
            Do While idx1 < Len(s)
                idx2 = InStr(idx1, s, Chr$(32))
                If (idx2 = ZERO) Then idx2 = Len(s) + ONE
                iLength = idx2 - idx1 ' truncates space char
                ReDim Preserve aStrsOut(ZERO To idxArr) As String
                aStrsOut(idxArr) = Mid$(s, idx1, iLength)
                idxArr = idxArr + ONE
                idx1 = idx2 + ONE
            Loop
        Else ' (iSpc <> ZERO) And (iQuot <> ZERO)
            idx1 = InStr(idx1, s, Chr$(34))
            Do While idx1 < Len(s)
                idx2 = InStr(idx1 + ONE, s, Chr$(34))
                If (idx2 = ZERO) Then Err.Raise INVALID_ARG, , "Invalid number of quotation characters."
                iLength = idx2 - idx1
                ReDim Preserve aStrsOut(ZERO To idxArr) As String
                aStrsOut(idxArr) = Mid$(s, idx1 + ONE, iLength - ONE)
                ' Strip quoted strings out of s
                If (idx1 = ONE) Then
                    s = Mid$(s, idx2 + ONE)
                Else
                    s = Mid$(s, ONE, idx1 - ONE) & Mid$(s, idx2 + ONE)
                End If
                idxArr = idxArr + ONE
                idx1 = InStr(idx1, s, Chr$(34))
                If (idx1 = ZERO) Then Exit Do
            Loop
            s = Trim$(s)
            ' Split non-quoted space delimited names
            Do While LenB(s) <> 0
                idx2 = InStr(ONE, s, Chr$(32))
                If (idx2 = ZERO) Then idx2 = Len(s) + ONE
                ReDim Preserve aStrsOut(ZERO To idxArr) As String
                aStrsOut(idxArr) = Mid$(s, ONE, idx2 - ONE)
                If (idx2 > Len(s)) Then Exit Do
                s = Mid$(s, idx2 + ONE)
                s = Trim$(s)
                idxArr = idxArr + ONE
            Loop
        End If
    End If
End Sub
