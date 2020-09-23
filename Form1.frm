VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "SendMessage API & 50 messages & samples & Comments"
   ClientHeight    =   5445
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6585
   LinkTopic       =   "Form1"
   ScaleHeight     =   5445
   ScaleWidth      =   6585
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "ReadMe1st"
      Height          =   615
      Left            =   120
      TabIndex        =   18
      Top             =   4560
      Width           =   1335
   End
   Begin VB.PictureBox Picture1 
      Height          =   855
      Index           =   4
      Left            =   2160
      ScaleHeight     =   795
      ScaleWidth      =   2235
      TabIndex        =   14
      Top             =   4560
      Width           =   2295
      Begin VB.TextBox Text1 
         Height          =   375
         Index           =   4
         Left            =   0
         TabIndex        =   15
         Text            =   "Text1"
         Top             =   360
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "The glory goes to:"
         Height          =   375
         Index           =   4
         Left            =   0
         TabIndex        =   16
         Top             =   120
         Width           =   1575
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   1095
      Index           =   3
      Left            =   0
      ScaleHeight     =   1035
      ScaleWidth      =   6435
      TabIndex        =   11
      Top             =   3360
      Width           =   6495
      Begin VB.TextBox Text1 
         Height          =   855
         Index           =   3
         Left            =   0
         MultiLine       =   -1  'True
         TabIndex        =   12
         Text            =   "Form1.frx":0000
         Top             =   240
         Width           =   6375
      End
      Begin VB.Label Label1 
         Caption         =   "Comment"
         Height          =   375
         Index           =   3
         Left            =   0
         TabIndex        =   13
         Top             =   0
         Width           =   1575
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   1095
      Index           =   2
      Left            =   4320
      ScaleHeight     =   1035
      ScaleWidth      =   2115
      TabIndex        =   8
      Top             =   2280
      Width           =   2175
      Begin VB.TextBox Text1 
         Height          =   495
         Index           =   2
         Left            =   0
         TabIndex        =   9
         Text            =   "Text1"
         Top             =   480
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "lParam"
         Height          =   375
         Index           =   2
         Left            =   240
         TabIndex        =   10
         Top             =   0
         Width           =   1575
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   1095
      Index           =   1
      Left            =   2160
      ScaleHeight     =   1035
      ScaleWidth      =   2115
      TabIndex        =   5
      Top             =   2280
      Width           =   2175
      Begin VB.TextBox Text1 
         Height          =   495
         Index           =   1
         Left            =   0
         TabIndex        =   6
         Text            =   "Text1"
         Top             =   480
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "wParam"
         Height          =   375
         Index           =   1
         Left            =   0
         TabIndex        =   7
         Top             =   0
         Width           =   1575
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   1095
      Index           =   0
      Left            =   0
      ScaleHeight     =   1035
      ScaleWidth      =   2115
      TabIndex        =   2
      Top             =   2280
      Width           =   2175
      Begin VB.TextBox Text1 
         Height          =   495
         Index           =   0
         Left            =   0
         TabIndex        =   4
         Text            =   "Text1"
         Top             =   480
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "Mesasage"
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   3
         Top             =   0
         Width           =   1575
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Save new record"
      Height          =   855
      Left            =   120
      TabIndex        =   1
      Top             =   1320
      Width           =   2775
   End
   Begin VB.ListBox List1 
      Height          =   2205
      Left            =   3000
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   120
      Width           =   3495
   End
   Begin VB.Label Label2 
      Caption         =   "Select items in list, syntax is automaticaly added to clipboard"
      Height          =   615
      Left            =   360
      TabIndex        =   17
      Top             =   120
      Width           =   2535
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim dbs As Database
Dim rst As Recordset
Sub refreshlist()
 rst.MoveFirst
 Do
 List1.AddItem rst![Message]
 rst.MoveNext
 Loop Until rst.EOF
End Sub
Private Sub Command1_Click()
MsgBox "Huh, it took me a lot of time to figure this all out, I would appreciate if you would share your dbtest.mdb with me. Idea is that we all together fill this list of send message examples to full extent possible. Of course I will publish new versions on PSC as sun as I get new samples. Wote if you like, but be sure to send me ur dbtest.mdb to kozlicki@yahoo.com. And yeah hex values of constants are not included, mybe in next version, look for API-Guide and Wiever on net to get them "




 

End Sub

Private Sub Command2_Click()
rst.MoveFirst
'check if message is already in database
Do
 If rst![Message] = Text1(0).Text Then
 a = MsgBox("This message is already in database, saving this will simply add new recors, be sure that you are adding something new. Do you want to save new record ?", vbYesNo, Note)
 End If
 rst.MoveNext
Loop Until rst.EOF
   Select Case a
   Case 6 'yes
   With rst
        .AddNew
        ![Message] = Text1(0).Text
        ![lParam] = Text1(1).Text
        ![wParam] = Text1(2).Text
        ![Comment] = Text1(3).Text
        ![Author] = Text1(4).Text
   End With
   refreshlist
   Case 7 'no
   Case Else
   End Select

End Sub

Private Sub Form_Load()
MsgBox "If u want to run this you need reference to MS DAO 3.51 or higher in your project"
 'password oppening, to show you a sample
Set dbs = OpenDatabase(App.Path & "\dbtest.mdb", False, False, ";pwd=M.C")
Set rst = dbs.OpenRecordset("DatabaseTable")
refreshlist
End Sub

Private Sub List1_Click()

rst.MoveFirst
 Do
 If rst![Message] = List1.List(List1.ListIndex) Then
Text1(0).Text = rst![Message]
Text1(1).Text = rst![wParam]
Text1(2).Text = rst![lParam]
Text1(3).Text = rst![Comment]
Text1(4).Text = rst![Author]
'add all this to clipboard
Clipboard.Clear
Clipboard.SetText "variable = SendMessage(control.hwnd," & Text1(0).Text & "," & Text1(1).Text & "," & Text1(2).Text & ")"
 Exit Do
 End If
 
 rst.MoveNext
 Loop Until rst.EOF
End Sub
