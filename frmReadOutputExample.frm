VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "Richtx32.ocx"
Begin VB.Form frmReadOutputExample 
   BackColor       =   &H00000000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Endra's MS-DOS [The Lost Version]"
   ClientHeight    =   3870
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   10605
   Icon            =   "frmReadOutputExample.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3870
   ScaleWidth      =   10605
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   9120
      MousePointer    =   1  'Arrow
      TabIndex        =   3
      Top             =   120
      Width           =   1335
   End
   Begin prjReadOutput.ReadOutput ReadOutput1 
      Left            =   9240
      Top             =   3000
      _ExtentX        =   1720
      _ExtentY        =   1296
   End
   Begin VB.CommandButton cmdExecute 
      Caption         =   "E&xecute"
      Height          =   375
      Left            =   7680
      MousePointer    =   1  'Arrow
      TabIndex        =   1
      Top             =   120
      Width           =   1335
   End
   Begin VB.TextBox txtCommand 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   405
      Left            =   4200
      TabIndex        =   0
      Text            =   "ping www.google.com"
      Top             =   120
      Width           =   3375
   End
   Begin RichTextLib.RichTextBox rtfOutput 
      Height          =   3135
      Left            =   120
      TabIndex        =   4
      Top             =   600
      Width           =   10335
      _ExtentX        =   18230
      _ExtentY        =   5530
      _Version        =   393217
      BackColor       =   0
      BorderStyle     =   0
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"frmReadOutputExample.frx":030A
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Dialog"
         Size            =   6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      BackColor       =   &H80000008&
      Caption         =   "Command to get output from:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   270
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   3915
   End
End
Attribute VB_Name = "frmReadOutputExample"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'You may use this code in your project as long as you dont claim its yours ;)

'NEW IN V2.1:
'   -Nice more DOS like environment
'   -Replaced TextBox with Rich Text Format Box (More then 65535 chars allowed)
'   -Added support to the DOS 'CLS' command (Clear Screen)
'   -Replaced output of DOS 'CD' command to be C:\
'   -Made it show default path as C:\>
'   -Colors changed
'   -Added KeyPress event so you can press ENTER instead of clicking Execute
'Thats about it..
'NO changes have been done in the control!
'Enjoy!

Option Explicit

Private Sub cmdCancel_Click()
    txtCommand.SetFocus
    ReadOutput1.CancelProcess
End Sub

Private Sub cmdExecute_Click()
    txtCommand.SetFocus
    ReadOutput1.SetCommand = txtCommand.Text
    ReadOutput1.ProcessCommand
End Sub

Private Sub Form_Load()
    Me.Show
    rtfOutput.Text = "[**] Endra's Version of MS-DOS [**]" & vbNewLine & vbNewLine & "C:\> "
    txtCommand.SetFocus
    rtfOutput.SelStart = 0
    rtfOutput.SelLength = Len(rtfOutput.Text)
    rtfOutput.SelColor = &H80000000
End Sub

Private Sub ReadOutput1_Canceled()
    rtfOutput.Text = rtfOutput.Text & vbNewLine & "[**] Process Canceled [**]" & vbNewLine & vbNewLine & "C:\> "
    MsgBox "Success! Process was canceled!"
End Sub

Private Sub ReadOutput1_Complete()
    If rtfOutput.Text = "" Then
        rtfOutput.Text = "C:\> "
    Else
        rtfOutput.Text = rtfOutput.Text & vbNewLine & "C:\> "
        MsgBox "Complete reading output!", vbOKOnly, "Success!"
    End If
End Sub

Private Sub ReadOutput1_Error(ByVal Error As String, LastDLLError As Long)
    MsgBox "Error!" & vbNewLine & _
            "Description: " & Error & vbNewLine & _
            "LastDLLError: " & LastDLLError, vbCritical, "Error"
End Sub

Private Sub ReadOutput1_GotChunk(ByVal sChunk As String, ByVal LastChunk As Boolean)
    'your probly wondering why I put LastChunk when I already put the Complete event..
    'if you test you'll see that you get chunk by chunk (256 chars), not line by line
    'so if you want to parse those, you'll need to know when it finishes so you can
    'release your last line since you cannot check if its complete by using the event.
    'LastChunk is false if there is more chunks, true if that is the last chunk.
    If Len(sChunk) >= 3 Then
        If Left(sChunk, Len(sChunk) - 2) = App.Path Then
            rtfOutput.Text = rtfOutput.Text & "C:\" & vbNewLine
            Exit Sub
        End If
    End If
    rtfOutput.Text = rtfOutput.Text & Replace(Replace(sChunk, Chr(13), ""), Chr(10), vbNewLine)
    If Len(sChunk) = 1 Then
        If Asc(sChunk) = 12 Then
            rtfOutput.Text = "[**] Endra's Version of MS-DOS [**]" & vbNewLine
        End If
    End If
End Sub

Private Sub ReadOutput1_Starting()
    rtfOutput.Text = rtfOutput.Text & ReadOutput1.SetCommand & vbNewLine
End Sub

Private Sub txtCommand_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        cmdExecute_Click
        KeyAscii = 0
    End If
End Sub

Private Sub rtfOutput_Change()
    rtfOutput.SelStart = Len(rtfOutput.Text) + 1
    rtfOutput.SelLength = Len(rtfOutput.Text) + 1
End Sub
