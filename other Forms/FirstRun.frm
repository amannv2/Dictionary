VERSION 5.00
Begin VB.Form FirstRun 
   BackColor       =   &H80000007&
   BorderStyle     =   0  'None
   ClientHeight    =   3570
   ClientLeft      =   10125
   ClientTop       =   2130
   ClientWidth     =   6180
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3570
   ScaleWidth      =   6180
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox Text2 
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
      Left            =   2880
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   1440
      Width           =   2775
   End
   Begin VB.TextBox Text1 
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
      Left            =   2880
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   720
      Width           =   2775
   End
   Begin VB.Image Image2 
      Height          =   765
      Left            =   3480
      Picture         =   "FirstRun.frx":0000
      Stretch         =   -1  'True
      Top             =   2400
      Width           =   2175
   End
   Begin VB.Image Image1 
      Height          =   735
      Left            =   480
      Picture         =   "FirstRun.frx":A49E
      Stretch         =   -1  'True
      Top             =   2400
      Width           =   2055
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Confirm Password"
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
      Left            =   240
      TabIndex        =   3
      Top             =   1440
      Width           =   2340
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
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
      Left            =   240
      TabIndex        =   1
      Top             =   720
      Width           =   1260
   End
End
Attribute VB_Name = "FirstRun"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim btn As Integer

Private Sub Image1_Click()
Unload Me
btn = 0
End Sub

Private Sub Image2_Click()
btn = 1
If Text1 <> "" Or Text2 <> "" Then
    If Text1 = Text2 Then
        Open App.Path & "\Admin_pass" For Output As #2
        Print #2, Text1.Text
        Close #2
        MsgBox "Password Successfully Applied"
            Unload Me
    
    Else
        MsgBox "Please Confirm password... By typing exactly same password on both boxes"
    End If
End If
End Sub
