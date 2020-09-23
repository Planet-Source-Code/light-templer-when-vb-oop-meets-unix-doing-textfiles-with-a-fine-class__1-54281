VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Simple demo to   'clsParseTextFile.clc'"
   ClientHeight    =   3660
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5175
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3660
   ScaleWidth      =   5175
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnGo 
      BackColor       =   &H00FAC5AD&
      Caption         =   "Do the job and output to debug window !"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1005
      Left            =   195
      MaskColor       =   &H00FFFFFF&
      Picture         =   "frmMain.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   390
      Width           =   4770
   End
   Begin VB.Label lblUpdate 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Update 2"
      Height          =   255
      Left            =   4275
      TabIndex        =   5
      Top             =   3225
      Width           =   720
   End
   Begin VB.Label lblHeader 
      BackStyle       =   0  'Transparent
      Caption         =   "If you like doing things with fine class ..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   225
      TabIndex        =   4
      Top             =   75
      Width           =   3915
   End
   Begin VB.Label lblDesc 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   $"frmMain.frx":1194
      Height          =   660
      Index           =   1
      Left            =   420
      TabIndex        =   3
      Top             =   1515
      Width           =   4260
   End
   Begin VB.Image imgLogo 
      Height          =   765
      Left            =   1260
      Picture         =   "frmMain.frx":1237
      Top             =   2760
      Width           =   2550
   End
   Begin VB.Label lblDesc 
      BackStyle       =   0  'Transparent
      Caption         =   "Created by Light Templer (LiTe) in June '04."
      Height          =   285
      Index           =   0
      Left            =   1050
      TabIndex        =   1
      Top             =   2385
      Width           =   3075
   End
   Begin VB.Label lblAuthor 
      BackStyle       =   0  'Transparent
      Caption         =   "Light Templer"
      ForeColor       =   &H0080FF80&
      Height          =   225
      Left            =   1875
      TabIndex        =   2
      Top             =   2370
      Width           =   990
   End
   Begin VB.Shape shpBorder1 
      Height          =   330
      Left            =   825
      Shape           =   4  'Rounded Rectangle
      Top             =   2325
      Width           =   3480
   End
   Begin VB.Shape shpBckGrnd1 
      BorderColor     =   &H00FAC5AD&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00FAC5AD&
      FillStyle       =   0  'Solid
      Height          =   195
      Left            =   825
      Top             =   2400
      Width           =   3480
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents oTEXT As clsParseTextFile
Attribute oTEXT.VB_VarHelpID = -1
'
'
'

Private Sub btnGo_Click()
    
    Dim sPathFilename   As String
    
    Set oTEXT = New clsParseTextFile
    With oTEXT
        
        .IgnoreEmptyLines = True    ' we don't need empty lines
        .IgnoreLinesWith = "'"      ' we don't want comment lines
        
        
        ' Example for CAT: Show ALL lines
        Debug.Print vbCrLf + "=== CAT " + String$(40, "=") + vbCrLf + vbCrLf
        If .CAT("D:\Projekte\ALL CLASS\clsParseTextFile\testfile.txt") = True Then
            MsgBox "Ready / Success!", vbInformation, " CAT"
        Else
            MsgBox "Ready / Error!" + vbCrLf + .LastErrMsg, vbExclamation, " CAT"
        End If
        Debug.Print vbCrLf
        Debug.Print "Total lines:  " & .LinesTotal & ",  handled lines:  " & .LinesToHandle & vbCrLf & vbCrLf



        ' Example for HEAD: Show 10 first line lines only
        Debug.Print vbCrLf + "=== HEAD " + String$(40, "=") + vbCrLf + vbCrLf
        If .HEAD(10, "D:\Projekte\ALL CLASS\clsParseTextFile\testfile.txt") = True Then
            MsgBox "Ready / Success!", vbInformation, " HEAD"
        Else
            MsgBox "Ready / Error!" + vbCrLf + .LastErrMsg, vbExclamation, " HEAD"
        End If
        Debug.Print "Total lines:  " & .LinesTotal & ",  handled lines:  " & .LinesToHandle & vbCrLf & vbCrLf



        ' Example for TAIL: Show last 8 lines only
        Debug.Print vbCrLf + "=== TAIL " + String$(40, "=") + vbCrLf + vbCrLf
        If .TAIL(8, "D:\Projekte\ALL CLASS\clsParseTextFile\testfile.txt") = True Then
            MsgBox "Ready / Success!", vbInformation, " TAIL"
        Else
            MsgBox "Ready / Error!" + vbCrLf + .LastErrMsg, vbExclamation, " TAIL"
        End If
        Debug.Print "Total lines:  " & .LinesTotal & ",  handled lines:  " & .LinesToHandle & vbCrLf & vbCrLf



        ' Example for FILTER: Show matching lines only (We use CAT here, but works with all commands!)
        Debug.Print vbCrLf + "=== FILTER " + String$(40, "=") + vbCrLf + vbCrLf
        .Filter = "*Public*"
        If .CAT("D:\Projekte\ALL CLASS\clsParseTextFile\testfile.txt") = True Then
            MsgBox "Ready / Success!", vbInformation, " FILTER"
        Else
            MsgBox "Ready / Error!" + vbCrLf + .LastErrMsg, vbExclamation, " FILTER"
        End If
        Debug.Print "Total lines:  " & .LinesTotal & ",  handled lines:  " & .LinesToHandle & vbCrLf & vbCrLf
        
        
        
        ' Example for GetTemp:  Get a new, unique filename in Windows temp dir
        Debug.Print vbCrLf + "=== GetTemp " + String$(40, "=") + vbCrLf
        sPathFilename = .GetTemp()
        ' sPathFilename = .GetTemp("C:\temp", "!MyTemp~")       ' with all parameters used
        If sPathFilename <> "" Then
            MsgBox "Ready / Success!" + vbCrLf + vbCrLf + sPathFilename, vbInformation, " GetTemp"
        Else
            MsgBox "Ready / Error!" + vbCrLf + .LastErrMsg, vbExclamation, " GetTemp"
        End If
        Debug.Print .GetTemp
        
        
        
        ' Example for Append:  Append a few lines to a file
        Debug.Print vbCrLf + "=== Append " + String$(40, "=") + vbCrLf
        If .Append("This is line 1", sPathFilename) = False Then
            MsgBox "Ready / Error!" + vbCrLf + .LastErrMsg, vbExclamation, " Append"
        End If
        If .Append("This is next line", sPathFilename) = False Then
            MsgBox "Ready / Error!" + vbCrLf + .LastErrMsg, vbExclamation, " Append"
        End If
        
        .Append "Two more lines without error checking"     ' In shortest form we leave the 2nd parameter empty,
        .Append "Next line without error checking"          ' so the last used filename for output is used again!
        
        If .Append("This is the last line!") = False Then
            MsgBox "Ready / Error!" + vbCrLf + .LastErrMsg, vbExclamation, " Append"
        End If
        MsgBox "Ready / Success!" + vbCrLf + vbCrLf + "Appended 3 lines to " + sPathFilename, vbInformation, " Append"
        Debug.Print sPathFilename + "has 3 more lines at the end!"
        
        
    End With
    Set oTEXT = Nothing
    
End Sub

Private Sub oTEXT_Error(sErrMsg As String, lLineNo As Long)
    
    MsgBox sErrMsg & " - Line# " & lLineNo, vbExclamation, "Error !"

    
End Sub

Private Sub oTEXT_HandleLine(lLineNo As Long, sLine As String)

    Debug.Print lLineNo, sLine
    
    ' if x = y then                         example for abort of parsing
    '     oTEXT.CancelParsing = True
    ' End If
    
End Sub

' #*#
