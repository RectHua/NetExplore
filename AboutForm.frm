VERSION 5.00
Begin VB.Form AboutForm 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "���� NetExplore �����"
   ClientHeight    =   4350
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8295
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4350
   ScaleWidth      =   8295
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.CommandButton Command1 
      Caption         =   "ȷ��"
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
      Left            =   6360
      TabIndex        =   0
      Top             =   3480
      Width           =   1455
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "�汾 2.0"
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   15
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   405
      Left            =   6240
      TabIndex        =   6
      Top             =   2280
      Width           =   1125
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "��Ȩ���� INTRON �����֯����������Ȩ����"
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   480
      TabIndex        =   5
      Top             =   3480
      Width           =   4095
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000010&
      X1              =   0
      X2              =   8280
      Y1              =   2880
      Y2              =   2880
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Explore"
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   36
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   855
      Left            =   2880
      TabIndex        =   4
      Top             =   1440
      Width           =   2535
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "INTRON �����֯ NetExplore IE ���ݲ�� NeXT"
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   480
      TabIndex        =   3
      Top             =   3240
      Width           =   4815
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "NeXT"
      BeginProperty Font 
         Name            =   "΢���ź� Light"
         Size            =   36
         Charset         =   134
         Weight          =   290
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   930
      Left            =   5640
      TabIndex        =   2
      Top             =   1440
      Width           =   1785
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "NetWork"
      BeginProperty Font 
         Name            =   "΢���ź� Light"
         Size            =   36
         Charset         =   134
         Weight          =   290
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   855
      Left            =   2880
      TabIndex        =   1
      Top             =   480
      Width           =   4455
   End
   Begin VB.Image Image1 
      Height          =   1920
      Left            =   480
      Picture         =   "AboutForm.frx":0000
      ToolTipText     =   "�¾�������ʱ������ - Internet Explorer"
      Top             =   480
      Width           =   1920
   End
End
Attribute VB_Name = "AboutForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()            '���ȷ����ť
    Unload Me
End Sub

Private Sub Form_Load()                 '���ڼ����¼�
Me.Icon = LoadPicture("")
End Sub

