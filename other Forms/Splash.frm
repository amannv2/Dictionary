VERSION 5.00
Begin VB.Form Splash 
   BackColor       =   &H00404040&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   4965
   ClientLeft      =   5265
   ClientTop       =   3795
   ClientWidth     =   8040
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   4965
   ScaleWidth      =   8040
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer10 
      Interval        =   50
      Left            =   1920
      Top             =   4440
   End
   Begin VB.Timer Timer9 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   7560
      Top             =   2760
   End
   Begin VB.Timer Timer8 
      Enabled         =   0   'False
      Interval        =   700
      Left            =   6480
      Top             =   4320
   End
   Begin VB.Timer Timer7 
      Enabled         =   0   'False
      Interval        =   700
      Left            =   6120
      Top             =   4320
   End
   Begin VB.Timer Timer6 
      Enabled         =   0   'False
      Interval        =   700
      Left            =   5760
      Top             =   4320
   End
   Begin VB.Timer Timer5 
      Interval        =   700
      Left            =   5400
      Top             =   4320
   End
   Begin VB.Timer Timer4 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   7560
      Top             =   1920
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   7560
      Top             =   1560
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   7560
      Top             =   1080
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   7560
      Top             =   600
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000C0&
      BorderWidth     =   5
      X1              =   0
      X2              =   360
      Y1              =   4920
      Y2              =   4920
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   3285
      TabIndex        =   11
      Top             =   4560
      Width           =   135
   End
   Begin VB.Label Label11 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Words Loaded"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   3480
      TabIndex        =   10
      Top             =   4560
      Width           =   1515
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "'s"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   600
      Left            =   4665
      TabIndex        =   9
      Top             =   2160
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Label Label9 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "ICTIONARY"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   600
      Left            =   5160
      TabIndex        =   8
      Top             =   2760
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   675
      Left            =   4920
      TabIndex        =   7
      Top             =   3840
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   675
      Left            =   4800
      TabIndex        =   6
      Top             =   3840
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   630
      Left            =   4800
      TabIndex        =   5
      Top             =   3840
      Visible         =   0   'False
      Width           =   255
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C0C0&
      Height          =   975
      Left            =   4230
      TabIndex        =   4
      Top             =   1320
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "D"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   72
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   1920
      Left            =   3150
      TabIndex        =   3
      Top             =   1320
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Loading"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   750
      Left            =   2535
      TabIndex        =   2
      Top             =   3840
      Width           =   1980
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "S"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   975
      Left            =   2685
      TabIndex        =   1
      Top             =   1200
      Visible         =   0   'False
      Width           =   405
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C0C0&
      Height          =   975
      Left            =   2070
      TabIndex        =   0
      Top             =   1155
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image Image1 
      Height          =   4935
      Left            =   0
      Picture         =   "Splash.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   8055
   End
End
Attribute VB_Name = "Splash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim i As Integer

Private Sub Form_Load()
i = 3

Label1.Top = 1400
Label2.Top = 1400
Label5.Top = 1400
Label4.Top = 1400
End Sub



Private Sub Timer1_Timer()
Label1.Top = Label1.Top + 10
Label1.Visible = True
If Label1.Top = 1800 Then Timer2.Enabled = True
If Label1.Top = 2000 Then Timer1.Enabled = False
End Sub



Private Sub Timer10_Timer()
Label12 = Label12.Caption + i
i = i + 10
If Val(Label12) >= 100000 Then
Label11.Visible = False
Label12 = "All Content Loaded"
Label12.Left = 2900
Timer10.Enabled = False
Load Main
Unload Me
On Error Resume Next
Main.Show
If Err Then Unload Me
    
End If
Line1.X2 = Line1.X2 + 55
End Sub

Private Sub Timer2_Timer()
Label2.Top = Label2.Top + 10
Label2.Visible = True
If Label2.Top = 1800 Then Timer3.Enabled = True
If Label2.Top = 2000 Then Timer2.Enabled = False
End Sub

Private Sub Timer3_Timer()
Label4.Top = Label4.Top + 10
Label4.Visible = True
If Label4.Top = 1800 Then Timer4.Enabled = True
Timer9.Enabled = True
If Label4.Top = 1800 Then Timer3.Enabled = False
End Sub

Private Sub Timer4_Timer()
Label5.Top = Label5.Top + 10
Label5.Visible = True
If Label5.Top = 2000 Then
Timer4.Enabled = False
End If
'If Label4.Top = 1800 Then Timer4.Enabled = True
End Sub

Private Sub Timer5_Timer()
Label6.Visible = True
Timer6.Enabled = True
Timer5.Enabled = False
End Sub

Private Sub Timer6_Timer()
Label7.Visible = True
Timer7.Enabled = True
Timer6.Enabled = False
End Sub

Private Sub Timer7_Timer()
Label8.Visible = True
Timer8.Enabled = True
Timer7.Enabled = False
End Sub

Private Sub Timer8_Timer()
Label6.Visible = False
Label7.Visible = False
Label8.Visible = False
Timer5.Enabled = True
Timer8.Enabled = False
End Sub

Private Sub Timer9_Timer()
Label9.Visible = True
Label9.Left = Label9.Left - 10
If Label9.Left = 4100 Then
Label10.Visible = True
Timer9.Enabled = False
End If
End Sub
