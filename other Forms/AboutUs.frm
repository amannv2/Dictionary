VERSION 5.00
Begin VB.Form AboutUs 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   6315
   ClientLeft      =   1185
   ClientTop       =   1425
   ClientWidth     =   8760
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6315
   ScaleWidth      =   8760
   ShowInTaskbar   =   0   'False
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Label4"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   555
      Left            =   0
      TabIndex        =   3
      Top             =   3720
      Width           =   8775
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   795
      Left            =   0
      TabIndex        =   2
      Top             =   1680
      Width           =   8775
   End
   Begin VB.Image Image2 
      Height          =   375
      Left            =   8400
      Picture         =   "AboutUs.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   2115
      Left            =   0
      TabIndex        =   1
      Top             =   4200
      Width           =   8775
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1395
      Left            =   0
      TabIndex        =   0
      Top             =   2400
      Width           =   8775
   End
   Begin VB.Image Image1 
      Height          =   1695
      Left            =   0
      Picture         =   "AboutUs.frx":D2A2
      Stretch         =   -1  'True
      Top             =   0
      Width           =   8775
   End
End
Attribute VB_Name = "AboutUs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Label3 = "A.S.D.A. Dictionary" & vbNewLine
Label1.Caption = "This software is protected under copyright laws. Copying or other illegal activity will lead " & "to Criminal Offence"
Label4.Caption = "Developed By..." & vbNewLine
Label2.Caption = "Aman Verma" & vbNewLine
 Label2.Caption = Label2.Caption + "Saurabh Singh" & vbNewLine
 Label2.Caption = Label2.Caption + "Divya Prakhar" & vbNewLine
 Label2.Caption = Label2.Caption + "Azizur Rehman"
 
                    
                    
                    
End Sub

Private Sub Image2_Click()
Unload Me
End Sub

