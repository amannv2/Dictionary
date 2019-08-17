VERSION 5.00
Begin VB.Form check 
   BackColor       =   &H80000007&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3750
   ClientLeft      =   9420
   ClientTop       =   2370
   ClientWidth     =   6660
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3750
   ScaleWidth      =   6660
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox passCheck 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   3360
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   840
      Width           =   2775
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Enter Password :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   330
      Left            =   720
      TabIndex        =   1
      Top             =   840
      Width           =   2160
   End
   Begin VB.Image Image1 
      Height          =   735
      Left            =   600
      Picture         =   "check.frx":0000
      Stretch         =   -1  'True
      Top             =   2160
      Width           =   2055
   End
   Begin VB.Image Image2 
      Height          =   765
      Left            =   3600
      Picture         =   "check.frx":3812
      Stretch         =   -1  'True
      Top             =   2160
      Width           =   2175
   End
End
Attribute VB_Name = "check"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public pass_String As String
Public btn As Integer

Private Sub Image1_Click()
Unload Me
btn = 0
End Sub

Private Sub Image2_Click()
btn = 1
If Len(passCheck) > 0 Then
pass_String = passCheck
Unload Me
Else
MsgBox "Please Enter a Valid Password", vbExclamation
End If
End Sub

