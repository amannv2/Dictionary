VERSION 5.00
Begin VB.Form Saved_words 
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   4260
   ClientLeft      =   9660
   ClientTop       =   1665
   ClientWidth     =   5520
   LinkTopic       =   "Form2"
   ScaleHeight     =   4260
   ScaleWidth      =   5520
   ShowInTaskbar   =   0   'False
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2655
      Left            =   240
      TabIndex        =   0
      Top             =   720
      Width           =   1815
   End
   Begin VB.Image Image2 
      Height          =   495
      Left            =   4920
      Picture         =   "Saved_words.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   615
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Kruti Dev 040 Wide"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1455
      Left            =   2640
      TabIndex        =   1
      Top             =   1200
      Width           =   2535
   End
   Begin VB.Image Image1 
      Height          =   4335
      Left            =   0
      Picture         =   "Saved_words.frx":D2A2
      Stretch         =   -1  'True
      Top             =   0
      Width           =   5535
   End
End
Attribute VB_Name = "Saved_words"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Image2_Click()
Unload Me
End Sub

Private Sub List1_Click()
        Main.REcSet.Open "Select * from English where English_word='" & List1.List(List1.ListIndex) & "'", Main.con, adOpenKeyset
          
          Label1 = Main.REcSet.Fields("Hindi_word")
If Main.REcSet.RecordCount = 0 Then
          Label1.Font = "Segoe UI"
          Label1.FontSize = 14
          Label1 = "This is your Added Word. It is currently unavailable in words list"
          Else
          Label1.Font = "Kruti Dev 010"
          Label1.FontSize = 26
End If
          
Main.REcSet.Close
End Sub
