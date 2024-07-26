VERSION 5.00
Begin VB.Form FavoriteForm 
   BackColor       =   &H8000000E&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "NeXT 收藏夹"
   ClientHeight    =   8010
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5940
   LinkTopic       =   "Form8"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8010
   ScaleWidth      =   5940
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "微软雅黑 Light"
         Size            =   10.5
         Charset         =   134
         Weight          =   290
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Text            =   " 在<这里>输入并按下回车键以自定义添加"
      Top             =   1920
      Width           =   5655
   End
   Begin VB.CommandButton Command4 
      Caption         =   "删除全部"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2880
      TabIndex        =   4
      Top             =   1440
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      Caption         =   "刷新记录"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4080
      TabIndex        =   3
      Top             =   1440
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "删除选中"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1680
      TabIndex        =   2
      Top             =   1440
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "添加当前页面"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   1440
      Width           =   1455
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5460
      Left            =   120
      TabIndex        =   0
      Top             =   2400
      Width           =   5655
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      BackStyle       =   0  'Transparent
      Caption         =   "NeXT 收藏夹"
      BeginProperty Font 
         Name            =   "微软雅黑 Light"
         Size            =   24
         Charset         =   134
         Weight          =   290
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   615
      Left            =   360
      TabIndex        =   6
      Top             =   360
      Width           =   2760
   End
   Begin VB.Image Image1 
      Height          =   1320
      Left            =   0
      Picture         =   "FavoriteForm.frx":0000
      Top             =   0
      Width           =   6615
   End
End
Attribute VB_Name = "FavoriteForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
On Error Resume Next
List1.RemoveItem List1.ListIndex

Open App.Path & "\favorites.dat" For Output As #25
For i = 0 To List1.ListCount - 1
Print #25, List1.List(i)
Next i
Close #25
    
End Sub

Private Sub Command2_Click()

Open App.Path & "\favorites.dat" For Append As #23
Print #23, MainForm.WebBrowser1.LocationURL
Close #23

List1.Clear
Dim Favorites2
Open App.Path & "\favorites.dat" For Input As #24
Do While Not EOF(24)
    Line Input #24, Favorites2
    List1.AddItem Favorites2
    Loop
    Close #24
    
End Sub

Private Sub Command3_Click()
List1.Clear
Dim Favorites3
Open App.Path & "\favorites.dat" For Input As #22
Do While Not EOF(22)
    Line Input #22, Favorites3
    List1.AddItem Favorites3
    Loop
    Close #22
    
End Sub

Private Sub Command4_Click()
Dim DELFAV As String
DELFAV = MsgBox("此操作不可挽回！确定要删除您的全部收藏吗？", vbOKCancel + vbInformation, "警告")
If DELFAV = vbOK And Not Dir(App.Path & "\favorites.dat") = "" Then
Kill App.Path & "\favorites.dat"
End If
Open App.Path & "\favorites.dat" For Append As #27
Close #27

List1.Clear
Dim Favorites6
Open App.Path & "\favorites.dat" For Input As #31
Do While Not EOF(31)
    Line Input #31, Favorites6
    List1.AddItem Favorites6
    Loop
    Close #31

End Sub

Private Sub Form_Load()
Me.Icon = LoadPicture("")

Dim Favorites
List1.Clear
Open App.Path & "\favorites.dat" For Append As #19
Close #19

Open App.Path & "\favorites.dat" For Input As #20
Do While Not EOF(20)
    Line Input #20, Favorites
    List1.AddItem Favorites
    Loop
    Close #20

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Text1.Text = " " Or Text1.Text = "" Then
Text1.Text = " 在<这里>输入并按下回车键以自定义添加"
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
Open App.Path & "\favorites.dat" For Output As #21
For i = 0 To List1.ListCount - 1
Print #21, List1.List(i)
Next i
Close #21
End Sub

Private Sub List1_DblClick()
On Error Resume Next
If Not List1.Text = "" Or List1.Text = " " Then
MainForm.WebBrowser1.Navigate List1.Text
Unload Me
End If
End Sub

Private Sub Text1_Click()
If Text1.Text = " 在<这里>输入并按下回车键以自定义添加" Then
Text1.Text = ""
End If
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)

If Text1.Text = " 在<这里>输入并按下回车键以自定义添加" Then
Text1.Text = ""
End If

If KeyAscii = 8 Then Exit Sub

If KeyAscii = 13 And Not Text1.Text = " " And Not Text1.Text = "" Then
List1.AddItem Text1.Text

Open App.Path & "\favorites.dat" For Output As #34
For i = 0 To List1.ListCount - 1
Print #34, List1.List(i)
Next i
Close #34

End If

End Sub
