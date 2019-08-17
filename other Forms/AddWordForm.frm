VERSION 5.00
Begin VB.Form AddWordForm 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   4110
   ClientLeft      =   9885
   ClientTop       =   2610
   ClientWidth     =   6615
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "AddWordForm.frx":0000
   ScaleHeight     =   4110
   ScaleWidth      =   6615
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox meaning 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Kruti Dev 010"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   2760
      TabIndex        =   2
      Top             =   1200
      Width           =   2775
   End
   Begin VB.TextBox word 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   2760
      TabIndex        =   0
      Top             =   480
      Width           =   2775
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Enter Meaning:"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   375
      Left            =   600
      TabIndex        =   3
      Top             =   1200
      Width           =   1920
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Enter Word :"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   375
      Left            =   600
      TabIndex        =   1
      Top             =   480
      Width           =   1575
   End
   Begin VB.Image Image1 
      Height          =   735
      Left            =   480
      Picture         =   "AddWordForm.frx":0DE4
      Stretch         =   -1  'True
      Top             =   2280
      Width           =   2055
   End
   Begin VB.Image Image2 
      Height          =   765
      Left            =   3360
      Picture         =   "AddWordForm.frx":45F6
      Stretch         =   -1  'True
      Top             =   2280
      Width           =   2175
   End
   Begin VB.Image Image3 
      Height          =   4095
      Left            =   0
      Picture         =   "AddWordForm.frx":EA94
      Stretch         =   -1  'True
      Top             =   0
      Width           =   6615
   End
End
Attribute VB_Name = "AddWordForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public word_string, meaning_string As String
Public btn_press As Integer

Private Sub Form_Load()
btn_press = 0
End Sub

Private Sub Image1_Click()
Unload Me
btn_press = 0
End Sub

Private Sub Image2_Click()

btn_press = 1
If Len(word) < 1 Or Len(meaning) < 1 Then
MsgBox "Please Enter Valid Information", vbExclamation
Else
meaning.Font = "Arial"
word_string = word
meaning_string = meaning
Unload Me

End If
End Sub

