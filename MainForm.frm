VERSION 5.00
Begin VB.Form Main 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "My Dictionary"
   ClientHeight    =   7785
   ClientLeft      =   840
   ClientTop       =   855
   ClientWidth     =   8355
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7785
   ScaleWidth      =   8355
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2160
      TabIndex        =   0
      Top             =   1800
      Width           =   2775
   End
   Begin VB.Timer view_setting 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   7440
      Top             =   1200
   End
   Begin VB.Timer SlideTimer 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   7440
      Top             =   840
   End
   Begin VB.Frame Drawer 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   7815
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   855
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Change Password"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   960
         TabIndex        =   6
         Top             =   4560
         Width           =   2535
      End
      Begin VB.Image icon 
         Height          =   735
         Index           =   4
         Left            =   0
         Picture         =   "MainForm.frx":0000
         Stretch         =   -1  'True
         Top             =   4440
         Width           =   735
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "About Us..."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   960
         TabIndex        =   5
         Top             =   5400
         Width           =   2175
      End
      Begin VB.Image Image5 
         Height          =   735
         Left            =   0
         Picture         =   "MainForm.frx":12882
         Stretch         =   -1  'True
         Top             =   5280
         Width           =   735
      End
      Begin VB.Image drawer_option 
         Height          =   375
         Index           =   4
         Left            =   1080
         Picture         =   "MainForm.frx":3160C
         Stretch         =   -1  'True
         Top             =   6480
         Width           =   1335
      End
      Begin VB.Image drawer_option 
         Height          =   375
         Index           =   3
         Left            =   960
         Picture         =   "MainForm.frx":321C3
         Stretch         =   -1  'True
         Top             =   3840
         Width           =   1815
      End
      Begin VB.Image drawer_option 
         Appearance      =   0  'Flat
         Height          =   375
         Index           =   2
         Left            =   960
         Picture         =   "MainForm.frx":336B9
         Stretch         =   -1  'True
         Top             =   2880
         Width           =   2535
      End
      Begin VB.Image drawer_option 
         Height          =   375
         Index           =   1
         Left            =   960
         Picture         =   "MainForm.frx":356FA
         Stretch         =   -1  'True
         Top             =   2040
         Width           =   2055
      End
      Begin VB.Image drawer_option 
         Height          =   375
         Index           =   0
         Left            =   960
         Picture         =   "MainForm.frx":371EB
         Stretch         =   -1  'True
         Top             =   1200
         Width           =   2415
      End
      Begin VB.Image Image1 
         Appearance      =   0  'Flat
         Height          =   615
         Left            =   0
         Picture         =   "MainForm.frx":39D28
         Stretch         =   -1  'True
         Top             =   6360
         Width           =   735
      End
      Begin VB.Image icon 
         Appearance      =   0  'Flat
         Height          =   615
         Index           =   3
         Left            =   0
         Picture         =   "MainForm.frx":3AC78
         Stretch         =   -1  'True
         Top             =   3720
         Width           =   735
      End
      Begin VB.Image icon 
         Appearance      =   0  'Flat
         Height          =   615
         Index           =   2
         Left            =   0
         Picture         =   "MainForm.frx":3BD05
         Stretch         =   -1  'True
         Top             =   2760
         Width           =   735
      End
      Begin VB.Image icon 
         Appearance      =   0  'Flat
         Height          =   495
         Index           =   1
         Left            =   0
         Picture         =   "MainForm.frx":3C898
         Stretch         =   -1  'True
         Top             =   2040
         Width           =   735
      End
      Begin VB.Image icon 
         Appearance      =   0  'Flat
         Height          =   615
         Index           =   0
         Left            =   0
         Picture         =   "MainForm.frx":3D710
         Stretch         =   -1  'True
         Top             =   1080
         Width           =   735
      End
      Begin VB.Image Drawer_btn 
         Height          =   495
         Left            =   0
         Picture         =   "MainForm.frx":3EBC2
         Stretch         =   -1  'True
         Top             =   240
         Width           =   615
      End
      Begin VB.Image drawer_bg 
         Height          =   7695
         Left            =   -120
         Picture         =   "MainForm.frx":4419C
         Stretch         =   -1  'True
         Top             =   0
         Width           =   3735
      End
   End
   Begin VB.ListBox List2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF8080&
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000004&
      Height          =   3180
      Left            =   2160
      Sorted          =   -1  'True
      TabIndex        =   3
      Top             =   2160
      Width           =   3135
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Search:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000004&
      Height          =   345
      Left            =   1080
      TabIndex        =   1
      Top             =   1800
      Width           =   1095
   End
   Begin VB.Image Image3 
      Height          =   615
      Left            =   7560
      Picture         =   "MainForm.frx":E30D6
      Stretch         =   -1  'True
      Top             =   0
      Width           =   735
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Kruti Dev 010"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   735
      Left            =   2160
      TabIndex        =   2
      Top             =   6360
      Width           =   3135
   End
   Begin VB.Image Image2 
      Height          =   375
      Left            =   4920
      Picture         =   "MainForm.frx":F69B8
      Stretch         =   -1  'True
      Top             =   1800
      Width           =   375
   End
   Begin VB.Image Image4 
      Height          =   1575
      Left            =   720
      Picture         =   "MainForm.frx":FCC73
      Stretch         =   -1  'True
      Top             =   0
      Width           =   7335
   End
   Begin VB.Image main_bg 
      Height          =   6375
      Left            =   720
      Picture         =   "MainForm.frx":10A9DF
      Stretch         =   -1  'True
      Top             =   1440
      Width           =   7335
   End
End
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public con As New ADODB.Connection
Dim str As String
Dim pos(1 To 5) As Integer
Public notify, c, X As Integer
Public REcSet As New ADODB.Recordset
Dim nrec As New ADODB.Recordset

Public Function check_pass(pstr As String) As Boolean
Dim pwd As String
On Error Resume Next
    Open App.Path & "\Admin_pass" For Input As #1
        Line Input #1, pwd          'Reading only the first line of the File
    Close #1
If Err Then
    MsgBox "You must be an Admin to perform this action! For Admin privilege, you must assign " & _
            "the password for this application... Thanks!", vbInformation
            check.btn = 0
            FirstRun.Show 1
Else
    If pstr = pwd Then
        check_pass = True
    Else
        check_pass = False
    End If
End If

End Function
Private Sub change_pass()

check.Show 1
If check_pass(check.pass_String) Then
FirstRun.Show 1
            
 ElseIf check.btn = 1 Then
 MsgBox "Wrong Password", vbCritical
 End If
 
End Sub

Private Sub form1_KeyDown(KeyCode As Integer, Shift As Integer)
X = List2.ListIndex
If List2.ListIndex > 0 And List2.ListIndex <= List2.ListCount And X > 0 Then
    If vbKeyDown Then
    X = X + 1
        List2.Selected(X) = True

    ElseIf vbKeyUp Then
    X = X - 1
        List2.Selected(X) = True
    End If
End If
End Sub

Private Sub Add_Word()
Dim insert_string As String

check.Show 1
If check_pass(check.pass_String) Then
        AddWordForm.Show 1
    If AddWordForm.btn_press = 1 Then
        word = AddWordForm.word_string
        meaning = AddWordForm.meaning_string
        
        insert_string = "insert into English(English_word,Hindi_word) values ('" & word & "','" & meaning & "')"
        On Error Resume Next
          con.Execute insert_string

            If Err Then MsgBox "Error in insertion"
            MsgBox "Word Successfulley Added!", vbInformation
            
    End If
            
ElseIf check.btn = 1 Then
    MsgBox "Password is NOT correct", vbExclamation
End If
End Sub

Private Sub Reset_DB()
Dim t As Integer
check.Show 1

       
If check_pass(check.pass_String) Then
    ans = MsgBox("Are You Sure you want to reset?", vbYesNo)
    If ans = vbYes Then
            con.Close
        On Error Resume Next
        Kill App.Path & "\dictionary DB.mdb"
            notify = 1
        FileCopy App.Path & "\DB Instance\dictionary DB.mdb", App.Path & "\dictionary DB.mdb"
        Text1 = ""
        List2.Clear
        Label2 = ""
        
        con.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & _
App.Path & "\dictionary DB.mdb;Jet OLEDB:Database Password=dont waste your time;"

        
            If Err Then
           MsgBox "Cannot reset"
            Else
                MsgBox "Successfully Reset", vbInformation
                notify = 0
            End If
    End If
ElseIf check.btn = 1 Then
    MsgBox "Password Incorrect!", vbExclamation
End If
End Sub

Private Sub Save_Favourite()
Dim sv_word As String
    If List2.ListCount >= 1 Then
        sv_word = List2.List(List2.ListIndex)
        
       Open App.Path & "\favrt " For Append As #1
                Print #1, sv_word
            Close #1
        MsgBox "Word Saved Successfully!!", vbInformation
                    
    End If
            
End Sub

Private Sub Show_favourite()
Dim l_word() As String
Dim word As String
On Error Resume Next
    Open App.Path & "\favrt" For Input As #1
        word = Input(LOF(1), #1)
            l_word = Split(word, vbNewLine)
    Close #1
    Debug.Print word
    If Err Then
    MsgBox "Favourite List is Empty!", vbInformation
    Else
        For i = 0 To UBound(l_word)
            Saved_words.List1.AddItem l_word(i)
        Next
        Saved_words.List1.RemoveItem (Saved_words.List1.ListCount - 1)
        Saved_words.List1.Selected(0) = True

        Saved_words.Show 1
    End If
End Sub




Private Sub Drawer_btn_Click()
SlideTimer.Enabled = True
If X = 0 Then
    X = 1
    
Else
    X = 0
        
End If
  
End Sub
Public Sub move_controls(a As Integer, s As Integer)
If a = 0 Then
    s = -s
Else
s = s
End If
    Image2.Left = Image2.Left + s
    Label1.Left = Label1.Left + s ' + 10
    Label2.Left = Label2.Left + s
    Text1.Left = Text1.Left + s
    List2.Left = List2.Left + s
    
End Sub

Private Sub drawer_option_Click(Index As Integer)
Select Case Index
    Case 0: Save_Favourite
    Case 1: Show_favourite
    Case 2: Add_Word
    Case 3: Reset_DB
    Case 4: Image3_Click
End Select
End Sub

Private Sub Form_Load()

Unload Splash
For p = 0 To 4
Label3.Visible = False
Label4.Visible = False
    
    drawer_option(p).Visible = False
Next


Drawer.Width = Drawer_btn.Width + 150
c = 0
X = 0
notify = 0
con.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & _
App.Path & "\dictionary DB.mdb;Jet OLEDB:Database Password=dont waste your time;"

On Error Resume Next                                'Checking if file exists
    Open App.Path & "\Admin_pass" For Input As #1
        If Err Then
            Open App.Path & "\Admin_pass" For Output As #2
                Close #2
        End If
    Line Input #1, fdata
    Close #1
Debug.Print fdata
    If Len(fdata) = 0 Then FirstRun.Show 1

End Sub

Private Sub Form_Unload(Cancel As Integer)
con.Close
End Sub




Private Sub icon_Click(Index As Integer)
Select Case Index
    Case 0: Save_Favourite
    Case 1: Show_favourite
    Case 2: Add_Word
    Case 3: Reset_DB
    Case 4: change_pass
End Select
End Sub

Private Sub Image1_Click()
Image3_Click
End Sub

Private Sub Image3_Click()
ans = MsgBox("Do You Really want to Exit?", vbYesNoCancel, "Exit")
If ans = vbYes Then End
End Sub




Private Sub Image5_Click()
AboutUs.Show 1

End Sub

Private Sub Label3_Click()
AboutUs.Show 1
End Sub

Private Sub Label4_Click()
change_pass
End Sub

Private Sub List2_Click()

If notify = 0 Then
 REcSet.Open "Select * from English where English_word ='" & List2.List(List2.ListIndex) & "'", con, adOpenKeyset
       Label2 = REcSet.Fields("Hindi_word")
    REcSet.Close
End If
End Sub



Private Sub SlideTimer_Timer()
Dim sh As Integer
sh = 200
If X = 0 Then
Drawer.Width = Drawer.Width - sh
    If Drawer.Width <= Drawer_btn.Width + 300 Then
    SlideTimer.Enabled = False
      view_setting.Enabled = False
      For p = 0 To 4
    drawer_option(p).Visible = False
    Label3.Visible = False
    Label4.Visible = False
    
        Next
      End If
  
    
ElseIf X = 1 Then
Drawer.Width = Drawer.Width + sh

    If Drawer.Width >= Me.Width / 2 - 800 Then
    SlideTimer.Enabled = False
        view_setting.Enabled = True
        End If
      
End If
move_controls X, sh
    
End Sub

Private Sub Text1_Change()
    Dim ind As Long
    
 If Len(Text1) <> 0 Then
            
        REcSet.Open "Select * from English where English_word like'" & Text1 & "%'", con, adOpenKeyset
     
     'ind = REcSet.Fields("Sno").Value
        
    '    Debug.Print List2.ListCount, ind, REcSet.RecordCount
           Label2 = REcSet.Fields("Hindi_word")
    
        List2.Clear
        
        
           For i = 1 To REcSet.RecordCount
            List2.AddItem REcSet.Fields("English_word")
            REcSet.MoveNext
            Next
          
          REc = REcSet.RecordCount
          REcSet.Close
End If

   If REc = 0 And Len(Text1) > 0 Then
   Label2 = "NO RECORD FOUND"
   Label2.Font = "Segoe UI"
      Label2.FontSize = 14

   Else
   Label2.Font = "Kruti Dev 010"
  If List2.ListCount > 0 Then List2.Selected(0) = True
   Label2.FontSize = 18
   End If
   
   If Len(Text1) = 0 Then
   List2.Clear
    Label2.Caption = ""
    End If

End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)

If KeyAscii >= 65 And KeyAscii <= 122 Or KeyAscii = 32 Or KeyAscii = 8 Then
KeyAscii = KeyAscii
Else
KeyAscii = 0
End If
End Sub

Private Sub view_setting_Timer()
c = c + 1
Debug.Print c
    If c > 10 And c < 20 Then
    drawer_option(0).Visible = True

    
    ElseIf c > 20 And c < 30 Then
    
    drawer_option(1).Visible = True
    
    
    ElseIf c > 30 And c < 40 Then
    
    drawer_option(2).Visible = True
    
    
    ElseIf c > 40 And c < 50 Then
    
    drawer_option(3).Visible = True
    
    ElseIf c > 50 And c < 60 Then
    drawer_option(4).Visible = True
    
    Label4.Visible = True
    ElseIf c > 60 Then
    Label3.Visible = True
    
    view_setting.Enabled = False
     c = 0
End If


End Sub
