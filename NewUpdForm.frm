VERSION 5.00
Begin VB.Form NewUpdForm 
   BackColor       =   &H8000000E&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   6375
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8655
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6375
   ScaleWidth      =   8655
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
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
      Height          =   615
      Left            =   6600
      TabIndex        =   8
      Top             =   5520
      Width           =   1695
   End
   Begin VB.Shape Shape1 
      Height          =   3735
      Left            =   360
      Top             =   1560
      Width           =   7935
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "4.����ڸ�DPI��ʾ���ϴ�NE����ģ���������ڿ���ͨ����Ӧ�ó���ļ�����ѡ�� -> ��DPI�޸� -> ʹ�á�ϵͳ����ǿ����ģʽ�������"
      BeginProperty Font 
         Name            =   "΢���ź� Light"
         Size            =   12
         Charset         =   134
         Weight          =   290
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   630
      Left            =   480
      TabIndex        =   7
      Top             =   4560
      Width           =   7650
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "3.�޸���һЩ BUG��������ÿ�θ����°汾���и�����ʾ�ˡ�"
      BeginProperty Font 
         Name            =   "΢���ź� Light"
         Size            =   12
         Charset         =   134
         Weight          =   290
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   480
      TabIndex        =   6
      Top             =   3960
      Width           =   6510
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "2.��һЩ�ط�����������޸ģ��������ٲ��ֹ��ܣ��������������ҳʱ������ʾ��ҳ�����ֱ����ˡ�"
      BeginProperty Font 
         Name            =   "΢���ź� Light"
         Size            =   12
         Charset         =   134
         Weight          =   290
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   630
      Left            =   480
      TabIndex        =   5
      Top             =   3000
      Width           =   7635
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1.������IE�������F12������ģʽ���ܣ����ڿ�����NE��ʹ��F12�ˡ�"
      BeginProperty Font 
         Name            =   "΢���ź� Light"
         Size            =   12
         Charset         =   134
         Weight          =   290
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   480
      TabIndex        =   4
      Top             =   2400
      Width           =   7350
   End
   Begin VB.Label Label3 
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
      ForeColor       =   &H8000000B&
      Height          =   405
      Left            =   7080
      TabIndex        =   3
      Top             =   720
      Width           =   1125
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
      ForeColor       =   &H8000000E&
      Height          =   930
      Left            =   4440
      TabIndex        =   2
      Top             =   240
      Width           =   1785
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "NetExplore"
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   36
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   855
      Left            =   360
      TabIndex        =   1
      Top             =   240
      Width           =   3855
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "���汾 NeXT ���¹���"
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   390
      Left            =   480
      TabIndex        =   0
      Top             =   1680
      Width           =   2955
   End
   Begin VB.Image Image2 
      Height          =   1320
      Left            =   6480
      Picture         =   "NewUpdForm.frx":0000
      Top             =   0
      Width           =   6615
   End
   Begin VB.Image Image1 
      Height          =   1320
      Left            =   0
      Picture         =   "NewUpdForm.frx":1C762
      Top             =   0
      Width           =   6615
   End
End
Attribute VB_Name = "NewUpdForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Form_Load()                 '���ڼ����¼�
Me.Icon = LoadPicture("")
End Sub

