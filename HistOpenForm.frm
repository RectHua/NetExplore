VERSION 5.00
Begin VB.Form HistOpenForm 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "�鿴��ʷ��¼"
   ClientHeight    =   3270
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4560
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3270
   ScaleWidth      =   4560
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   Begin VB.CommandButton Command2 
      Caption         =   "�ر�"
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3000
      TabIndex        =   3
      Top             =   2520
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "����"
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1320
      TabIndex        =   2
      Top             =   2520
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   960
      Width           =   4095
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "ʱ�䣺"
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   360
      TabIndex        =   4
      Top             =   480
      Width           =   630
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "���⣺"
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   360
      TabIndex        =   1
      Top             =   120
      Width           =   630
   End
End
Attribute VB_Name = "HistOpenForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
MainForm.WebBrowser1.Navigate Text1.Text
Unload HistOpenForm
Unload HistoryForm
End Sub

Private Sub Command2_Click()
Unload HistOpenForm
End Sub

Private Sub Form_Load()
Me.Icon = LoadPicture("")
Text1.Text = HistoryForm.URLList.List(HistoryForm.MainList.ListIndex)
Label1.Caption = "���⣺" & HistoryForm.TitleList.List(HistoryForm.MainList.ListIndex)
Label2.Caption = "ʱ�䣺" & HistoryForm.TimeList.List(HistoryForm.MainList.ListIndex)
End Sub

