VERSION 5.00
Begin VB.Form ControlForm 
   BackColor       =   &H8000000B&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "设置 NeXT"
   ClientHeight    =   6495
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5925
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6495
   ScaleWidth      =   5925
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "控制面板"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   3840
      TabIndex        =   11
      Top             =   5520
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "保存设置"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   3840
      TabIndex        =   5
      Top             =   4440
      Width           =   1455
   End
   Begin VB.ComboBox Combo1 
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
      Left            =   600
      TabIndex        =   4
      Text            =   "Internet Explorer 11"
      Top             =   3840
      Width           =   4695
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H8000000B&
      BorderStyle     =   0  'None
      Caption         =   " 设置浏览器主页 "
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   240
      TabIndex        =   0
      Top             =   1680
      Width           =   5415
      Begin VB.CommandButton Command4 
         BackColor       =   &H00FFFFFF&
         Caption         =   "保存设置"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   3600
         TabIndex        =   2
         Top             =   240
         Width           =   1455
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   1
         Text            =   "http://"
         Top             =   840
         Width           =   4695
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         BackStyle       =   0  'Transparent
         Caption         =   " 设置浏览器主页 "
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   0
         Width           =   1380
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "当 NetExplore 启动时会自动访问它。"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   360
         TabIndex        =   3
         Top             =   360
         Width           =   3135
      End
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "转到控制面板以对 IE 进行具体设置。"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   600
      TabIndex        =   12
      Top             =   5640
      Width           =   2985
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      BackStyle       =   0  'Transparent
      Caption         =   " INTERNET 设置 "
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   360
      TabIndex        =   10
      Top             =   5160
      Width           =   1470
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      BackStyle       =   0  'Transparent
      Caption         =   "NeXT 控制中心"
      BeginProperty Font 
         Name            =   "微软雅黑 Light"
         Size            =   24
         Charset         =   134
         Weight          =   290
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   615
      Left            =   360
      TabIndex        =   9
      Top             =   360
      Width           =   3240
   End
   Begin VB.Image Image1 
      Height          =   1320
      Left            =   0
      Picture         =   "ControlForm.frx":0000
      Top             =   0
      Width           =   6615
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      BackStyle       =   0  'Transparent
      Caption         =   " 设置IE内核 "
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   360
      TabIndex        =   7
      Top             =   3360
      Width           =   1005
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "NE 会使用指定的 IE 内核渲染页面。"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   600
      TabIndex        =   6
      Top             =   4560
      Width           =   2940
   End
End
Attribute VB_Name = "ControlForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Combo1.Text = "Internet Explorer 11" Then Call PrintInf(App.Path & "\neset.ini", "Trident", "11000")
If Combo1.Text = "Internet Explorer 10" Then Call PrintInf(App.Path & "\neset.ini", "Trident", "10000")
If Combo1.Text = "Internet Explorer 9" Then Call PrintInf(App.Path & "\neset.ini", "Trident", "9000")
If Combo1.Text = "Internet Explorer 8" Then Call PrintInf(App.Path & "\neset.ini", "Trident", "8000")
If Combo1.Text = "Internet Explorer 7" Then Call PrintInf(App.Path & "\neset.ini", "Trident", "7000")

End Sub

Private Sub Command2_Click()
Shell "control.exe inetcpl.cpl", 1                          '调用IE自带设置
End Sub

Private Sub Command4_Click()

Text1.Text = Replace(Text1.Text, " ", "")

If InStr(Text1.Text, "http://") = 0 And InStr(Text1.Text, "https://") = 0 Then
    Text1.Text = "http://" & Text1.Text
End If

Call PrintInf(App.Path & "\neset.ini", "Home", Text1.Text)

End Sub


Private Sub Form_Load()
    
    Me.Icon = LoadPicture("")

Text1.Text = ReadInf(App.Path & "\neset.ini", "Home")

With Combo1
    .AddItem "Internet Explorer 11"
    .AddItem "Internet Explorer 10"
    .AddItem "Internet Explorer 9"
    .AddItem "Internet Explorer 8"
    .AddItem "Internet Explorer 7"
End With

If ReadInf(App.Path & "\neset.ini", "Trident") = "11000" Then Combo1.Text = "Internet Explorer 11"
If ReadInf(App.Path & "\neset.ini", "Trident") = "10000" Then Combo1.Text = "Internet Explorer 10"
If ReadInf(App.Path & "\neset.ini", "Trident") = "9000" Then Combo1.Text = "Internet Explorer 9"
If ReadInf(App.Path & "\neset.ini", "Trident") = "8000" Then Combo1.Text = "Internet Explorer 8"
If ReadInf(App.Path & "\neset.ini", "Trident") = "7000" Then Combo1.Text = "Internet Explorer 7"

End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 32 Then KeyAscii = 0
If KeyAscii = 13 Then

Text1.Text = Replace(Text1.Text, " ", "")

If InStr(Text1.Text, "http://") = 0 And InStr(Text1.Text, "https://") = 0 Then
    Text1.Text = "http://" & Text1.Text
End If

Call PrintInf(App.Path & "\neset.ini", "Home", Text1.Text)

End If

End Sub
