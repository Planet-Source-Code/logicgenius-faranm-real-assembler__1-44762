VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Assembler >> Compile DOS Executable .com files"
   ClientHeight    =   5955
   ClientLeft      =   2460
   ClientTop       =   1920
   ClientWidth     =   6870
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5955
   ScaleWidth      =   6870
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   240
      MousePointer    =   3  'I-Beam
      MultiLine       =   -1  'True
      TabIndex        =   8
      Top             =   1440
      Width           =   1860
   End
   Begin VB.Timer tmrErr 
      Enabled         =   0   'False
      Interval        =   4000
      Left            =   1440
      Top             =   2400
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Height          =   1455
      Left            =   2400
      Locked          =   -1  'True
      MousePointer    =   1  'Arrow
      MultiLine       =   -1  'True
      TabIndex        =   17
      Top             =   1440
      Width           =   2055
   End
   Begin VB.TextBox txtApp 
      BackColor       =   &H00008000&
      ForeColor       =   &H0000FF00&
      Height          =   285
      Left            =   2760
      TabIndex        =   0
      Text            =   "Hello"
      Top             =   3180
      Width           =   3855
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Full Compile using MS Dos Debugger"
      Height          =   495
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   5040
      Width           =   2895
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Debug &App"
      Height          =   495
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   5040
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Save &Debug Script"
      Height          =   495
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5040
      Width           =   1575
   End
   Begin VB.TextBox txtPath 
      BackColor       =   &H00008000&
      ForeColor       =   &H0000FF00&
      Height          =   285
      Left            =   2760
      TabIndex        =   1
      Text            =   "C:\Windows\Desktop"
      Top             =   3480
      Width           =   3855
   End
   Begin VB.CommandButton command1 
      Caption         =   $"frmMain.frx":08CA
      Default         =   -1  'True
      Height          =   1425
      Left            =   4680
      MaskColor       =   &H00000000&
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1440
      Width           =   1860
   End
   Begin VB.Frame Frame1 
      Caption         =   "Debug  using DOS debugger"
      ForeColor       =   &H8000000D&
      Height          =   1815
      Left            =   360
      TabIndex        =   10
      Top             =   3960
      Width           =   6375
      Begin VB.TextBox txtAddress 
         BackColor       =   &H00008000&
         ForeColor       =   &H0000FF00&
         Height          =   285
         Left            =   2640
         TabIndex        =   13
         Text            =   "0100"
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox txtCX 
         BackColor       =   &H00008000&
         ForeColor       =   &H0000FF00&
         Height          =   285
         Left            =   2640
         TabIndex        =   12
         Text            =   "15"
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Starting Address in Memory :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   195
         Left            =   240
         TabIndex        =   15
         Top             =   360
         Width           =   2085
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Exe size in bytes (Hex Value) :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   195
         Left            =   240
         TabIndex        =   14
         Top             =   600
         Width           =   2190
      End
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "?"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   270
      Left            =   1920
      TabIndex        =   18
      ToolTipText     =   "Help"
      Top             =   1080
      Width           =   135
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "DOS Output :"
      ForeColor       =   &H8000000D&
      Height          =   195
      Left            =   2400
      TabIndex        =   16
      Top             =   1080
      Width           =   960
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "App name :"
      ForeColor       =   &H8000000D&
      Height          =   195
      Left            =   360
      TabIndex        =   11
      Top             =   3240
      Width           =   810
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Save File in :"
      ForeColor       =   &H8000000D&
      Height          =   195
      Left            =   360
      TabIndex        =   9
      Top             =   3480
      Width           =   915
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Source Program :"
      ForeColor       =   &H8000000D&
      Height          =   195
      Left            =   240
      TabIndex        =   7
      Top             =   1080
      Width           =   1230
   End
   Begin VB.Label Label1 
      BackColor       =   &H00008000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Quick Start >> Here is a valid Assembly program below. Click Assemble to compile the program to a DOS com file."
      ForeColor       =   &H0000FF00&
      Height          =   675
      Left            =   240
      TabIndex        =   6
      Top             =   240
      Width           =   6285
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text

Const AH = 180
Const AL = 176
Const AX = 184
Const BX = 187
Const CX = 185
Const DX = 186

Private program(40) As String
Private errLine As Long

Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function TerminateProcess Lib "kernel32" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long

Private Sub Command1_Click()

On Error Resume Next

Dim n As Integer, nCounter As Integer, nMen As Long
Dim chCmd As String, chBuffer As String, strHex As String
Dim hHandle As Long

 tmrErr.Enabled = True

 Kill txtPath.Text & "\" & txtApp.Text & ".com"
 Open txtPath.Text & "\" & txtApp.Text & ".com" For Binary As #1

    For n = LBound(program) To UBound(program)
       program(n) = ""
    Next n
    
    For n = 1 To Len(Text1.Text)
       If Mid$(Text1.Text, n, 1) = vbCr Then nCounter = nCounter + 1: n = n + 2
       program(nCounter) = program(nCounter) & Mid$(Text1.Text, n, 1)
    Next n

    For n = LBound(program) To UBound(program)
       
       If Len(program(n)) <> 0 Then chCmd = program(n) Else _
         Exit For
         errLine = n
         
         If InStr(chCmd, "mov") <> 0 Then
             nCounter = InStr(chCmd, "h")
             
             If nCounter <> 0 Then
                nCounter = nCounter - 1
             Else
                nCounter = InStr(chCmd, "x")
                If nCounter <> 0 Then
                   nCounter = nCounter - 1
                Else
                   nCounter = InStr(chCmd, "l")
                   nCounter = nCounter - 1
                End If
             End If
             
             nMen = Mnem(Mid$(chCmd, nCounter, 2))
             Do
                nCounter = nCounter + 1
             Loop Until IsNumeric(Mid$(chCmd, nCounter, 2)) = True Or _
               nCounter = Len(chCmd)

               strHex = Mid$(chCmd, nCounter + 1)
              
                 If Len(strHex) <= 2 Then
                    strHex = Hexer(strHex)
                    
                    Put #1, , Chr$(nMen)
                    Put #1, , Chr$(strHex)
                    
                    If nMen <> AH And nMen <> DX And nMen <> AL Then Put #1, , Chr$(0)
                    
                 Else
                 
                     If Len(strHex) = 3 Then strHex = "0" & strHex
                      Put #1, , Chr$(nMen)
                      Put #1, , Chr$(Val(Hexer(Right$(strHex, 2))))
                      Put #1, , Chr$(Val(Hexer(Left$(strHex, 2))))
                     If nMen <> AH And nMen <> DX And nMen <> AL Then Put #1, , Chr$(0)
                     
                 End If
         Else
         
           If InStr(chCmd, "int") <> 0 Then
            nCounter = 0
            
             Do
                nCounter = nCounter + 1
             Loop Until IsNumeric(Mid$(chCmd, nCounter, 2))
           
             nCounter = Val(Mid$(chCmd, nCounter + 1))
             
             Put #1, , Chr$(205)
             Put #1, , Chr$(nCounter + 12)
            
           Else
           
             If InStr(chCmd, "db") <> 0 Then
             
               Dim x As Integer, y As Integer
               Dim chStr As String, st As String
               Dim f As Boolean
                 
                 chStr = "": st = "": y = 1
                                  
                For nCounter = InStr(chCmd, "db") + 3 To Len(chCmd)
                   DoEvents
                   
                   If Mid$(chCmd, nCounter, 1) = "'" Then
                     If st <> "" Then chStr = chStr & Chr$(Val(st))
                      
                      st = ""
                      f = True
                   End If
                   
                  If f = False Then
                   If Mid$(chCmd, nCounter, 1) <> " " Then
                     If Len(st) <> 2 Then
                      st = st & CStr(Mid$(chCmd, nCounter, 1))
                     End If
                   End If
                   
                  Else
                   y = InStr(nCounter + 1, chCmd, "'")
                   chStr = chStr & Mid$(chCmd, nCounter + 1, y - 1 - nCounter)
                   f = False
                   nCounter = y + 1
                  End If
                   
                Next nCounter
                 
                If st <> "" Then chStr = chStr & Chr$(Val(st))
                 
                Put #1, , chStr
                
             End If
           End If
       End If
         
         
    Next n

Close #1
  
  Open txtPath.Text & "\tmp.bat" For Output As #1
    Print #1, txtApp.Text & ".com >tmpout"
  Close #1

  hHandle = Shell(txtPath.Text & "\tmp.bat")
  hHandle = OpenProcess(&H400, 1, (hHandle))
      
     Do
     Loop Until Dir(txtPath.Text & "\tmpout") <> ""
      
     For nCounter = 1 To 900000: Next
      
  TerminateProcess hHandle, &O0
  Kill txtPath.Text & "\tmp.bat"
 
  Open txtPath.Text & "\tmpout" For Binary As #1
    chBuffer = Space$(LOF(1))
    Get #1, , chBuffer
  Close #1

  Text2.Text = chBuffer
  Kill txtPath.Text & "\tmpout"
  tmrErr.Enabled = False
  Shell txtPath.Text & "\" & txtApp.Text & ".com", vbNormalFocus
End Sub

Private Sub Command2_Click()
  Dim n As Long, nCounter As Long
      
    For n = LBound(program) To UBound(program)
       program(n) = ""
    Next n
    
    For n = 1 To Len(Text1.Text)
       If Mid$(Text1.Text, n, 1) = vbCr Then nCounter = nCounter + 1: n = n + 2
       program(nCounter) = program(nCounter) & Mid$(Text1.Text, n, 1)
    Next n
      
   Open txtPath.Text & "\" & txtApp.Text & ".scr" For Output As #1
       
     Print #1, "a " & CStr(Val(txtAddress.Text))
       
     For n = LBound(program) To UBound(program)
        If Len(program(n)) <> 0 Then Print #1, program(n)
     Next n
       
     Print #1, vbCrLf & "r cx"
     Print #1, txtCX.Text
     Print #1, "n " & txtApp.Text & ".com"
     Print #1, "w"
     Print #1, "q" & vbCrLf
     
   Close #1
   
   Shell "notepad.exe " & txtPath.Text & "\" & txtApp.Text & ".scr", vbNormalFocus
End Sub

Private Sub Command3_Click()
   If Dir(txtPath.Text & "\" & txtApp.Text & ".com") = "" Then MsgBox "Compile First": Exit Sub
   
   Open Environ$("windir") & "\tmp.bat" For Output As #1
     Print #1, "@echo off"
     Print #1, Left(Environ$("windir"), 2)
     Print #1, "cd\"
     Print #1, "cd " & Mid$(Environ$("windir"), 4)
     Print #1, "@debug.exe " & txtApp.Text & ".com"
     Print #1, "@exit"
   Close #1
       
   Shell Environ$("windir") & "\tmp.bat", vbNormalFocus
End Sub

Private Sub Command4_Click()
   Dim n As Long, nCounter As Long
      
    For n = LBound(program) To UBound(program)
       program(n) = ""
    Next n
    
    For n = 1 To Len(Text1.Text)
       If Mid$(Text1.Text, n, 1) = vbCr Then nCounter = nCounter + 1: n = n + 2
       program(nCounter) = program(nCounter) & Mid$(Text1.Text, n, 1)
    Next n
      
   Open txtPath.Text & "\" & txtApp.Text & ".scr" For Output As #1
       
     Print #1, "a " & CStr(Val(txtAddress.Text))
       
     For n = LBound(program) To UBound(program)
        If Len(program(n)) <> 0 Then Print #1, program(n)
     Next n
       
     Print #1, vbCrLf & "r cx"
     Print #1, txtCX.Text
     Print #1, "n " & txtApp.Text & ".com"
     Print #1, "w"
     Print #1, "q" & vbCrLf
     
   Close #1
   
   Open txtPath.Text & "\" & txtApp.Text & ".bat" For Output As #1
     Print #1, "cls"
     Print #1, "@debug.exe <" & txtPath.Text & "\" & txtApp.Text & ".scr"
     Print #1, "@exit"
   Close #1
   
   Shell txtPath.Text & "\" & txtApp.Text & ".bat", vbMinimizedNoFocus
   MsgBox "Compiled using DEBUG.EXE. Created '" & txtApp.Text & "' batch and script files"
End Sub

Private Sub Label8_Click()
   Dim strOut As String
   
    strOut = "MOV AH,9           * Load AH register with no of DOS routine to use"
   strOut = strOut & vbCrLf & "MOV DX, 0109    * Load DX register with the address of string resource"
   strOut = strOut & vbCrLf & "INT 21                 * Call to DOS"
   strOut = strOut & vbCrLf & "INT 20                 * Return to DOS prompt"
   strOut = strOut & vbCrLf & "DB 'Hello' 32        * Define resource: Hello + Chr$(32)"
   strOut = strOut & vbCrLf & "DB 'World !' 07     *        + World + Chr$(7) or alert bell"
   strOut = strOut & vbCrLf & "DB '$'                   *  End resource"
   
   MsgBox strOut
   MsgBox "Statements allowed are: INT, MOV and DB" & vbCrLf & "Register allowed are: AH, AL, AX, BX, CX, DX"
   MsgBox "Compile program using MS DOS debugger for using full facilities nCounter.e. all statements, registers, etc"
End Sub

Private Sub Form_Load()
   Dim n As Integer
    
    program(0) = "MOV AH, 09"
    program(1) = "MOV DX, 0109"
    program(2) = "INT 21"
    program(3) = "INT 20"
    program(4) = "DB 'Hello' 32"
    program(5) = "DB 'World !' 07"
    program(6) = "DB '$'"
    
    For n = LBound(program) To UBound(program)
       If Len(program(n)) <> 0 Then _
         Text1.Text = Text1.Text & program(n) & vbCrLf
    Next n
    
    txtPath.Text = App.Path
End Sub

Function Hexer(no As String) As String
  If Len(no) <= 2 Then
     If Len(no) = 1 Then no = CStr(Val("0" & Trim(str$(no))))
     
     Hexer = CStr(Val("&h" & CStr(no)))
     Exit Function
  End If

End Function

Function Mnem(str As String) As Long
    
    Select Case str
       Case "ah"
        Mnem = AH
       Case "ax"
        Mnem = AX
       Case "bx"
        Mnem = BX
       Case "cx"
        Mnem = CX
       Case "dx"
        Mnem = DX
       Case "al"
        Mnem = AL
    End Select
    
End Function

Private Sub tmrErr_Timer()
    MsgBox "An unexpected exception occurred in in line no " & (errLine + 1) _
     & vbCrLf & "Line: " & program(errLine), vbCritical, "Error"
    End
End Sub
