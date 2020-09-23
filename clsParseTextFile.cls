VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsParseTextFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"

'
'   clsParseTextFile.cls
'

'   Created by:     Light Templer
'   Started   :     05/26/2004
'   Last edit :     06/15/2004

'   Purpose   :     Do some text file parsing in an object oriented, easy to use way.
'                   Even for very large files we don't need many memory. (This class
'                   doesn't read the whole file into memory, we are parsing
'                   line-by-line. For TAIL() a ring buffer is used.)
'
'
'                   So far this class encapsulates the following functions/properties:
'                   __________________________________________________________________
'                   Append()            Append a text to a file (file will be created if needed)
'                   CAT()               As Unix CAT command (gives ALL lines)
'                   HEAD()              As Unix HEAD command (gives first n lines)
'                   TAIL()              As Unix TAIL command (gives last n lines)
'                   GetTemp()           Get a unique temp path/filename
'
'                   CancelParsing       Abort parsing if set to true (e.g. in raised event)
'                   Filter              Filter lines (comparing with VB's LIKE)
'                   IgnoreEmptyLines    Ignore empty lines
'                   IgnoreLinesWith     Ignore lines starting with this value (e.g. to skip comment lines)
'                   LinesToHandle       Lines raised to deal with
'                   LinesTotal          Number of all the lines we read from text file
'                   LastErrMsg          Get last error message
'                   __________________________________________________________________
                                        
                                        
' Update 1    :     Thx for the hint to Sven Maes! If 'IgnoreLinesWith' was empty, ALL lines were skiped. Fixed.
' Update 2    :     This one I discovered by myself. The check for missing path/filename was against the wrong var.
                                        
                                        
' SHORT EXAMPLE:
'
'   Private WithEvents oTEXT As clsParseTextFile
'_________________________________________________
'
'   Private Sub main()
'
'      Set oTEXT = New clsParseTextFile
'      With oTEXT
'          .IgnoreEmptyLines = True
'          .IgnoreLinesWith = "'"
'          .Head 10, "C:\Temp\testfile.txt"         ' Show first 10 lines of this file
'          Debug.Print "Total lines:  " & .LinesTotal & ",  handled lines:  " & .LinesToHandle
'      End With
'      Set oTEXT = Nothing
'   End Sub
'
'   Private Sub oTEXT_Error(sErrMsg As String, lLineNo As Long)
'      Debug.Print "Error:  " & sErrMsg & " - Line# " & lLineNo
'   End Sub
'
'   Private Sub oTEXT_HandleLine(lLineNo As Long, sLine As String)
'      Debug.Print lLineNo, sLine
'   End Sub




Option Explicit
Option Compare Text                     ' Change to 'Option Compare Binary', if the 'Filter' property (VB 'Like'
                                        ' used in 'IsLineValid()' doesn't what you want!


' *******************************
' *           EVENTS            *
' *******************************
Public Event HandleLine(lLineNo As Long, sLine As String)
Public Event Error(sErrMsg As String, lLineNo As Long)



' *************************************
' *        API DEFINITIONS            *
' *************************************
Private Declare Function API_GetTempFileName Lib "kernel32" Alias "GetTempFileNameA" _
        (ByVal Path As String, _
         ByVal PrefixString As String, _
         ByVal Unique As Long, _
         ByVal TempFileName As String) As Long

Private Declare Function API_GetTempPath Lib "kernel32" Alias "GetTempPathA" _
        (ByVal nBufferLength As Long, _
         ByVal lpBuffer As String) As Long



' *************************************
' *        PRIVATE TYPES              *
' *************************************

' Here we hold all local info in a handy way
Private Type tpMvars
    sPathFilename       As String       ' Input
    sOutputFile         As String       ' Output (used with Append() only!)
    sErrMsg             As String       ' Last Error
    
    sIgnoreLinesWith    As String       ' e.g. = ";" -> we ignore lines when there first non-space char
                                        '  is ";" .  Used to ignore comments or so
                                        
    sRegExp             As String       ' For filtering lines with VB's 'LIKE' string compare function
    
    flgIgnoreEmptyLines As Boolean
    flgAbortParsing     As Boolean
    
    lLinesTotal         As Long
    lLinesToHandle      As Long
End Type
Private mvars As tpMvars
'
'
'



' *************************************
' *      PUBLIC SUBS/FUNCTIONS        *
' *************************************

Public Function Append(sText As String, Optional sPathFilename As String) As Boolean
    ' Append a string to a file. File will be created if it doesn't exists.
    ' Use it lie >> in DOS command shell/Unix shell.
    ' If 'sPathFilename' is empty last output file is used. If you never have used one, you 'll get an error!
    
    With mvars
    
        ' Check parameters
        If sPathFilename = "" Then
            If .sOutputFile = "" Then
                .sErrMsg = "Need a filename to append text in 'Append()'"
                raiseErr .sErrMsg
                    
                Exit Function
            End If
        Else
            .sOutputFile = sPathFilename
        End If
        
        If sText = "" Then          ' Nothing to do ... (for a new line use sText= vbCrLf !)
            ' Append = True         ' Maybe you don't this reported as an error: Remove comment sign.
            
            Exit Function
        End If
        
        ' Start appending
        If Do_APPEND(sText, .sOutputFile) = True And .sErrMsg = "" Then
            Append = True
        End If
    End With
    
End Function


Public Function CAT(Optional sPathFilename As String = "") As Boolean
    ' Returns True when parsing is done without any errors
    ' As the UNIX cataloge command CAT: This gives every single line of a text file, line by line
    
    ResetAll
    
    ' Check parameters
    If sPathFilename <> "" Then
        mvars.sPathFilename = sPathFilename
    End If
    
    If mvars.sPathFilename = "" Then
        mvars.sErrMsg = "Need a filename to parse in procedure 'Cat()'"
        mvars.flgAbortParsing = True
        raiseErr mvars.sErrMsg
                
        Exit Function
    End If
    
    ' Start parsing
    If Do_CAT() = True And mvars.sErrMsg = "" Then
        CAT = True
    End If
    
End Function


Public Function GetTemp(Optional sPath As String, Optional sPrefix As String = "~") As String
    ' Wrapper arround two API functions to get a new, unique filename in 'sPath'.
    ' If 'sPath' is empty, we ask Windows for current temp directory.
    ' If 'sPath' points to an existing directory, we get the unique filename within this directory
    ' 'sPrefix' is a leading part of this filename. If empty, the default ('~') will be used.
    '
    ' RESULT:   A valid path/filename to open a file, e.g.  "C:\Temp\~B7.tmp"
    ' On error: Result is empty and 'mvars.sErrMsg' holds the error message.


    Const CREATION_ERROR = 0
    Const MAX_PATH As Long = 260

    Dim sBuffer     As String
    Dim lResult     As Long
    
    
    ' Ensure having a path
    If sPath = "" Then
        sBuffer = Space$(MAX_PATH)
        lResult = API_GetTempPath(MAX_PATH, sBuffer)
        If lResult > 0 Then
            sPath = Left$(sBuffer, lResult)
        Else
            mvars.sErrMsg = "Error getting Windows Temp-directory by API!"
        
            Exit Function
        End If
    End If
    
    ' Get a new, unique filename
    sBuffer = Space$(MAX_PATH)
    If API_GetTempFileName(sPath, sPrefix, 0&, sBuffer) = CREATION_ERROR Then
        mvars.sErrMsg = "Error getting an unique filename in '" + sPath + "'by API!"
        
        Exit Function
    End If
    
    GetTemp = Left$(sBuffer, InStr(1, sBuffer, vbNullChar) - 1)
    
End Function


Public Function HEAD(NumberOfLines As Long, Optional sPathFilename As String = "") As Boolean
    ' Returns True when parsing is done without any errors
    ' As the UNIX command HEAD: This gives the first n lines of a text file, line by line
    
    ResetAll
    
    ' Check parameters
    If NumberOfLines < 1 Then
        mvars.sErrMsg = "Wrong parameter! Number of requested lines below 1 in procedure 'Head()'"
        mvars.flgAbortParsing = True
        raiseErr mvars.sErrMsg
                
        Exit Function
    End If
    
    If sPathFilename <> "" Then
        mvars.sPathFilename = sPathFilename
    End If
    
    If mvars.sPathFilename = "" Then
        mvars.sErrMsg = "Need a filename to parse in procedure 'Head()'"
        mvars.flgAbortParsing = True
        raiseErr mvars.sErrMsg
                
        Exit Function
    End If
    
    ' Start parsing
    If Do_HEAD(NumberOfLines) = True And mvars.sErrMsg = "" Then
        HEAD = True
    End If
    
End Function


Public Function TAIL(NumberOfLines As Long, Optional sPathFilename As String = "") As Boolean
    ' Returns True when parsing is done without any errors
    ' As the UNIX command TAIL: This gives the last n lines of a text file, line by line
    
    ResetAll
    
    ' Check parameters
    If NumberOfLines < 1 Then
        mvars.sErrMsg = "Wrong parameter! Number of requested lines below 1 in procedure 'Tail()'"
        mvars.flgAbortParsing = True
        raiseErr mvars.sErrMsg
                
        Exit Function
    End If
    
    If sPathFilename <> "" Then
        mvars.sPathFilename = sPathFilename
    End If
    
    If mvars.sPathFilename = "" Then
        mvars.sErrMsg = "Need a filename to parse in procedure 'Tail()'"
        mvars.flgAbortParsing = True
        raiseErr mvars.sErrMsg
                
        Exit Function
    End If
    
    ' Start parsing
    If Do_TAIL(NumberOfLines) = True And mvars.sErrMsg = "" Then
        TAIL = True
    End If
    
End Function



' *************************************
' *         PRIVATE FUNCTIONS         *
' *************************************
Private Function Do_APPEND(sText As String, sPathFilename As String) As Boolean
    ' Append a string to a file. File will be created if it doesn't exists.
    
    Dim lFhndl      As Long
        
    On Local Error GoTo error_handler
    
    
    lFhndl = FreeFile
    Open sPathFilename For Append As #lFhndl
    Print #lFhndl, sText
    Close #lFhndl
    
    Do_APPEND = True
    
    Exit Function


error_handler:
    
    mvars.sErrMsg = "[" & Err.Description & "] in procedure 'Do_Append()', text was: '" & sText & "'"
        
    If lFhndl Then
        Close #lFhndl
    End If
    raiseErr mvars.sErrMsg

End Function


Private Function Do_CAT() As Boolean
    ' Gives every single line of a text file, line by line
    
    Dim lFhndl      As Long
    Dim sLine       As String

    On Local Error GoTo error_handler


    lFhndl = FreeFile
    With mvars
        .flgAbortParsing = False
        Open .sPathFilename For Input As #lFhndl
        Do While EOF(lFhndl) = False And .flgAbortParsing = False
            Line Input #lFhndl, sLine
            .lLinesTotal = .lLinesTotal + 1
            If IsLineValid(sLine) = True Then
                .lLinesToHandle = .lLinesToHandle + 1
                
                ' Finally, here we give the line outside
                RaiseEvent HandleLine(.lLinesToHandle, sLine)
                
            End If
        Loop
        Close #lFhndl
    End With
    Do_CAT = True
    
    Exit Function


error_handler:
    
    mvars.sErrMsg = "[" & Err.Description & "] in procedure 'Do_CAT()', text line was: " & mvars.lLinesTotal
    mvars.flgAbortParsing = True
    
    If lFhndl Then
        Close #lFhndl
    End If
    raiseErr mvars.sErrMsg

End Function


Private Function Do_HEAD(lNumOfLines As Long) As Boolean
    ' Gives the first n lines of a text file, line by line

    Dim lFhndl      As Long
    Dim sLine       As String

    On Local Error GoTo error_handler


    lFhndl = FreeFile
    With mvars
        .flgAbortParsing = False
        Open .sPathFilename For Input As #lFhndl
        Do While EOF(lFhndl) = False And .flgAbortParsing = False
            Line Input #lFhndl, sLine
            .lLinesTotal = .lLinesTotal + 1
            If IsLineValid(sLine) = True Then
                .lLinesToHandle = .lLinesToHandle + 1
                ' Finally, here we give the line outside
                RaiseEvent HandleLine(.lLinesToHandle, sLine)
            End If
            If .lLinesToHandle >= lNumOfLines Then
                
                Exit Do
            End If
        Loop
        Close #lFhndl
    End With
    Do_HEAD = True
    
    Exit Function


error_handler:
    
    mvars.sErrMsg = "[" & Err.Description & "] in procedure 'Do_HEAD()', text line was: " & mvars.lLinesTotal
    mvars.flgAbortParsing = True
    
    If lFhndl Then
        Close #lFhndl
    End If
    raiseErr mvars.sErrMsg

End Function


Private Function Do_TAIL(lNumOfLines As Long) As Boolean
    ' Gives the last n lines of a text file, line by line
    '
    ' TACTIC:   Divided into two parts:
    '           1- We read the whole text file, line by line. To avoid wasting
    '              memory and to be able to parse very large files we only save
    '              the last 'lNumOfLines' lines read in a circular buffer.
    '           2- After this we write (raise) this buffer line by line.
    '
    ' HINT:     Filtering empty lines and line to ignore is done in part 1!
    '           So we really got 'lNumOfLines' valid lines (or less if file is too small)
    '           Because of this '.lLinesTotal' are always equal to '.lLinesToHandle' on TAIL!
    
    Dim lFhndl          As Long
    Dim sLine           As String
    Dim sLineBuffer()   As String       ' Here we hold the last 'lNumOfLines' lines in a circular buffer
    Dim lPtrIntoBuffer  As Long
    Dim i               As Long
    
    On Local Error GoTo error_handler

    ReDim sLineBuffer(1 To lNumOfLines)
    lFhndl = FreeFile
    With mvars
        .flgAbortParsing = False
        
        ' Part 1 - Read in and save into ring buffer
        Open .sPathFilename For Input As #lFhndl
        Do While EOF(lFhndl) = False
            Line Input #lFhndl, sLine
            If IsLineValid(sLine) = True Then
                
                ' New position in ring buffer (reaching the end means: start from beginning!)
                lPtrIntoBuffer = lPtrIntoBuffer + 1
                If lPtrIntoBuffer > lNumOfLines Then
                    lPtrIntoBuffer = 1
                End If
                
                ' Save a line a current position in ring buffer
                sLineBuffer(lPtrIntoBuffer) = sLine
                
                ' We need the true number of lines saved in buffer to get them out.
                ' Size of buffer is max for this.
                If .lLinesTotal < lNumOfLines Then
                    .lLinesTotal = .lLinesTotal + 1
                End If
                
            End If
        Loop
        Close #lFhndl
        
        ' Part 2 - Give the lines saved in ring buffer to caller
        If .lLinesTotal > 0 Then                                    ' If we really have lines ...
            lPtrIntoBuffer = IIf(.lLinesTotal < lNumOfLines, 0, lPtrIntoBuffer)
            Do
                i = i + 1
                lPtrIntoBuffer = lPtrIntoBuffer + 1
                If lPtrIntoBuffer > lNumOfLines Then
                    lPtrIntoBuffer = 1                              ' Jump to start when end is reached
                End If
                RaiseEvent HandleLine(i, sLineBuffer(lPtrIntoBuffer))
            Loop While i < .lLinesTotal And .flgAbortParsing = False
        End If
        .lLinesToHandle = .lLinesTotal
        
    End With
    Do_TAIL = True
    
    Exit Function


error_handler:
    
    mvars.sErrMsg = "[" & Err.Description & "] in procedure 'Do_TAIL()', text line was: " & mvars.lLinesTotal
    mvars.flgAbortParsing = True
    
    If lFhndl Then
        Close #lFhndl
    End If
    raiseErr mvars.sErrMsg

End Function


Private Function IsLineValid(sLine As String) As Boolean
    ' Here all checking for every line is done:  Use the line or skip it?
    
    Dim sLineTrimed     As String
   
    With mvars
        
        sLineTrimed = Trim$(sLine)
        
        ' Handle empty line
        If sLineTrimed = "" And .flgIgnoreEmptyLines = True Then
        
            Exit Function
        End If
        
        ' Handle ignoring
        If Len(.sIgnoreLinesWith) Then                      ' Changed (fixed) - thx to Sven Maes!
            If Left$(sLineTrimed, Len(.sIgnoreLinesWith)) = .sIgnoreLinesWith Then
            
                Exit Function
            End If
        End If
                
        ' Handle filtering with LIKE
        If Len(.sRegExp) Then
            If Not (sLine Like .sRegExp) Then
                
                Exit Function
            End If
        End If
        
    End With
    IsLineValid = True

End Function

Private Sub raiseErr(sErr As String)

    RaiseEvent Error(sErr, mvars.lLinesTotal)

End Sub


Private Sub ResetAll()
    ' Reset for a new start
    
    With mvars
        .flgAbortParsing = False
        .lLinesToHandle = 0
        .lLinesTotal = 0
        .sErrMsg = ""
    End With

End Sub




' *************************************
' *           PROPERTIES              *
' *************************************

Public Property Get LinesToHandle() As Long
    ' How many line were to handle (without ignored ones, e.g. when flgIgnoreEmptyLines= True !)
    
    LinesToHandle = mvars.lLinesToHandle
 
End Property

Public Property Get LinesTotal() As Long
    ' How many lines in textfile
    
    LinesTotal = mvars.lLinesTotal
 
End Property

Public Property Let CancelParsing(ByVal flgAbort As Boolean)
    ' Abort parsing when flag set to true (e.g. in raised event 'HandleLine()'
    
    mvars.flgAbortParsing = flgAbort
 
End Property

Public Property Get ParsingCanceled() As Boolean
    ' Was parsing aborted by user?
    
    ParsingCanceled = mvars.flgAbortParsing
 
End Property

Public Property Let IgnoreEmptyLines(ByVal flgIgnore As Boolean)
    ' We don't raise on empty lines
    
    mvars.flgIgnoreEmptyLines = flgIgnore
 
End Property

Public Property Get IgnoreEmptyLines() As Boolean
    ' For later questions or just to be complete in properties ;)
    
    IgnoreEmptyLines = mvars.flgIgnoreEmptyLines
 
End Property

Public Property Let IgnoreLinesWith(ByVal sStartString As String)
    ' We don't raise on lines starting with this string. e.g. used to ignore comment lines
    
    mvars.sIgnoreLinesWith = sStartString
 
End Property

Public Property Get IgnoreLinesWith() As String
    ' For later questions or just to be complete in properties ;)
    
    IgnoreLinesWith = mvars.sIgnoreLinesWith
 
End Property

Public Property Let Filter(ByVal sRegExp As String)
    ' We only raise on lines matching this 'regular expression' filter (in form of VB's 'LIKE' command)
    
    mvars.sRegExp = sRegExp
 
End Property

Public Property Get Filter() As String
    ' For later questions or just to be complete in properties ;)
    
    Filter = mvars.sRegExp
 
End Property

Public Property Get LastErrMsg() As String

    LastErrMsg = mvars.sErrMsg

End Property


Public Property Let PathFilename(ByVal sPathFilename As String)
    ' Set file here or give filename as arg on parse call
    
    mvars.sPathFilename = sPathFilename
 
End Property

Public Property Get PathFilename() As String
    ' File used
    
    PathFilename = mvars.sPathFilename
 
End Property

Public Property Let OutputFile(ByVal sPathFilename As String)
    ' Set file here or give filename as arg on Append() call
    
    mvars.sOutputFile = sPathFilename
 
End Property

Public Property Get OutputFile() As String
    ' File used for output with Append()
    
    OutputFile = mvars.sOutputFile
 
End Property


' #*#
