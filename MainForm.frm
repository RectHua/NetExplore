VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form MainForm 
   BackColor       =   &H00FFFFFF&
   Caption         =   " NetExplore"
   ClientHeight    =   10575
   ClientLeft      =   225
   ClientTop       =   270
   ClientWidth     =   16980
   Icon            =   "MainForm.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   10575
   ScaleWidth      =   16980
   StartUpPosition =   2  '��Ļ����
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   8520
      Top             =   120
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000006&
      Height          =   420
      Left            =   3840
      TabIndex        =   19
      Top             =   650
      Width           =   7575
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   420
      Left            =   2880
      Locked          =   -1  'True
      TabIndex        =   18
      Text            =   "��ȫ"
      Top             =   650
      Width           =   855
   End
   Begin VB.CommandButton Command5 
      Caption         =   "�ղ�"
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   14400
      TabIndex        =   17
      Top             =   120
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "��ʷ"
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   15240
      TabIndex        =   15
      Top             =   120
      Width           =   735
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   9375
      Left            =   0
      TabIndex        =   7
      Top             =   1215
      Width           =   16935
      ExtentX         =   29871
      ExtentY         =   16536
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   420
      Left            =   11520
      TabIndex        =   5
      Text            =   " ʹ�� Microsoft Bing ����"
      Top             =   650
      Width           =   5295
   End
   Begin VB.CommandButton Command4 
      Caption         =   "��"
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   2160
      TabIndex        =   4
      ToolTipText     =   "�����������ʱ��Ӧ��ӭ�Ѷ��ϣ������Ƿ������ҡ�"
      Top             =   650
      Width           =   615
   End
   Begin VB.CommandButton Command3 
      Caption         =   "��"
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   1440
      TabIndex        =   3
      ToolTipText     =   "��ʱ�򣬺���������ȷ��ѡ��"
      Top             =   650
      Width           =   615
   End
   Begin VB.CommandButton Command8 
      Caption         =   "��ҳ"
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   12720
      TabIndex        =   2
      Top             =   120
      Width           =   735
   End
   Begin VB.CommandButton Command6 
      Caption         =   "ѡ��"
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   16080
      TabIndex        =   1
      Top             =   120
      Width           =   735
   End
   Begin VB.CommandButton Command2 
      Caption         =   "ˢ��"
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   13560
      TabIndex        =   0
      Top             =   120
      Width           =   735
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "���ӵ���������Internet ��Ϊ�������ͬ�������ԣ�"
      BeginProperty Font 
         Name            =   "΢���ź� Light"
         Size            =   14.25
         Charset         =   134
         Weight          =   290
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1440
      TabIndex        =   16
      Top             =   2880
      Width           =   8055
   End
   Begin VB.Line Line1 
      BorderColor     =   &H8000000D&
      X1              =   0
      X2              =   16920
      Y1              =   1200
      Y2              =   1200
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "| ��������һ������·������������źŽ�������"
      BeginProperty Font 
         Name            =   "΢���ź� Light"
         Size            =   14.25
         Charset         =   134
         Weight          =   290
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1560
      TabIndex        =   14
      Top             =   7440
      Width           =   8055
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "| ������û�йرշ���ģʽ��"
      BeginProperty Font 
         Name            =   "΢���ź� Light"
         Size            =   14.25
         Charset         =   134
         Weight          =   290
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1560
      TabIndex        =   13
      Top             =   6840
      Width           =   8055
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "| ���ȷ�����������Ƿ񶼲���ˡ�"
      BeginProperty Font 
         Name            =   "΢���ź� Light"
         Size            =   14.25
         Charset         =   134
         Weight          =   290
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1560
      TabIndex        =   12
      Top             =   6240
      Width           =   8055
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "| ˢ�µ�ǰҳ��������·��ʡ�"
      BeginProperty Font 
         Name            =   "΢���ź� Light"
         Size            =   14.25
         Charset         =   134
         Weight          =   290
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1560
      TabIndex        =   11
      Top             =   5400
      Width           =   8055
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "| ȷ�� NetExplore û�б�ɱ������ͷ���ǽ���ء�"
      BeginProperty Font 
         Name            =   "΢���ź� Light"
         Size            =   14.25
         Charset         =   134
         Weight          =   290
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1560
      TabIndex        =   10
      Top             =   4800
      Width           =   8055
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "| ȷ��·�������������������ź��ȶ���"
      BeginProperty Font 
         Name            =   "΢���ź� Light"
         Size            =   14.25
         Charset         =   134
         Weight          =   290
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1560
      TabIndex        =   9
      Top             =   4200
      Width           =   8055
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "��... NetExplore �Ҳ��� Internet"
      BeginProperty Font 
         Name            =   "΢���ź� Light"
         Size            =   26.25
         Charset         =   134
         Weight          =   290
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   735
      Left            =   1440
      TabIndex        =   8
      Top             =   2040
      Width           =   9615
   End
   Begin VB.Image Image1 
      Height          =   960
      Left            =   240
      Picture         =   "MainForm.frx":4492
      ToolTipText     =   "�¾�������ʱ������ - Internet Explorer"
      Top             =   105
      Width           =   960
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "ʹ�� Trident ��Ⱦ���� - NetExplore 50"
      BeginProperty Font 
         Name            =   "΢���ź� Light"
         Size            =   15
         Charset         =   134
         Weight          =   290
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   1440
      TabIndex        =   6
      Top             =   120
      Width           =   5775
   End
   Begin VB.Menu control 
      Caption         =   "�������"
      Visible         =   0   'False
      Begin VB.Menu WYE 
         Caption         =   "��ҳ(&H)"
         Begin VB.Menu BACK 
            Caption         =   "����(&B)"
            Shortcut        =   ^B
         End
         Begin VB.Menu forward 
            Caption         =   "ǰ��(&F)"
            Shortcut        =   ^N
         End
         Begin VB.Menu stop 
            Caption         =   "ֹͣ(&O)"
            Shortcut        =   ^O
         End
         Begin VB.Menu fresh 
            Caption         =   "ˢ��(&R)"
            Shortcut        =   ^K
         End
         Begin VB.Menu home 
            Caption         =   "��ҳ(&H)"
            Shortcut        =   ^H
         End
      End
      Begin VB.Menu prin 
         Caption         =   "��ӡ(&P)"
         Begin VB.Menu print 
            Caption         =   "��ӡ(P)..."
            Shortcut        =   ^P
         End
      End
      Begin VB.Menu file 
         Caption         =   "�ļ�(&F)"
         Begin VB.Menu QuanScreen 
            Caption         =   "ȫ��(&L)      "
            Shortcut        =   {F11}
         End
         Begin VB.Menu LCW 
            Caption         =   "���Ϊ(&A)..."
            Shortcut        =   ^S
         End
         Begin VB.Menu SearchThing 
            Caption         =   "������ҳ����(S)..."
         End
      End
      Begin VB.Menu Display 
         Caption         =   "��ʾ(&D)"
         Begin VB.Menu ZoomSetUp 
            Caption         =   "�Ŵ� Ctrl +������ּӣ�"
         End
         Begin VB.Menu SmallSuoxiao 
            Caption         =   "��С Ctrl -������ּ���"
         End
         Begin VB.Menu ThereIsNotLine 
            Caption         =   "-"
         End
         Begin VB.Menu SF150 
            Caption         =   "���� 150%"
         End
         Begin VB.Menu SF100 
            Caption         =   "���� 100%"
         End
         Begin VB.Menu SF80 
            Caption         =   "���� 80%"
         End
         Begin VB.Menu SF50 
            Caption         =   "���� 50%"
         End
         Begin VB.Menu SF20 
            Caption         =   "���� 20%"
         End
         Begin VB.Menu NoneLineMenu 
            Caption         =   "-"
         End
         Begin VB.Menu SFZDY 
            Caption         =   "�Զ�������"
            Shortcut        =   ^Z
         End
      End
      Begin VB.Menu Nones 
         Caption         =   "-"
         Index           =   0
      End
      Begin VB.Menu Settings 
         Caption         =   "����(&S)"
      End
      Begin VB.Menu ModeSystem 
         Caption         =   "��ͼ(&M)"
         Begin VB.Menu GotoSmallWindow 
            Caption         =   "С��ģʽ(&S)"
         End
         Begin VB.Menu GotoNoHistory 
            Caption         =   "�����޺����(&H)"
         End
      End
      Begin VB.Menu TOOLS 
         Caption         =   "����(&T)"
         Begin VB.Menu F12open 
            Caption         =   "F12�����߹���(&L)"
            Shortcut        =   ^Y
         End
         Begin VB.Menu FavoritePage 
            Caption         =   "��ҳ�ղؼ�(&F)"
         End
         Begin VB.Menu HistoryForm1 
            Caption         =   "�����ʷ(&H)"
            Shortcut        =   ^U
         End
      End
      Begin VB.Menu None 
         Caption         =   "-"
      End
      Begin VB.Menu NewNE 
         Caption         =   "NE �������ܣ�&N��"
      End
      Begin VB.Menu about 
         Caption         =   "���� NetWork Explore (&A)"
         Shortcut        =   ^A
      End
   End
End
Attribute VB_Name = "MainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function InternetGetConnectedState Lib "wininet.dll" (ByRef dwFlags As Long, ByVal dwReserved As Long) As Long
Private Declare Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" (ByVal pCaller As Long, ByVal szURL As String, ByVal szFileName As String, ByVal dwReserved As Long, ByVal lpfnCB As Long) As Long
Private PrivateMode As String

Private Sub ABOUT_Click()           '����˵����еĹ���NE

    AboutForm.Show 1
    '������ǰ��ʾ���ڴ���
    
End Sub

Private Sub BACK_Click()            '����˵����еķ��ذ�ť

    On Error Resume Next
    WebBrowser1.GoBack
    
End Sub

Private Sub LCW_Click()            '����˵����е����Ϊѡ��

    WebBrowser1.ExecWB OLECMDID_SAVEAS, OLECMDEXECOPT_DODEFAULT
    
End Sub

Private Sub Command1_Click()            '�������ʷ����ť

    HistoryForm.Show 1

End Sub
Private Sub Command5_Click()                '������ղء���ť

    FavoriteForm.Show 1

End Sub

Private Sub FavoritePage_Click()                '����˵����еġ���ҳ�ղؼС�
    
    FavoriteForm.Show 1

End Sub

Private Sub GotoNoHistory_Click()               '����˵����е������޺�ģʽ

    ' ��������һ����Ϊ PrivateMode �� String �������Ѿ���ͨ�ò���������������Ϊ��¼�޺�ģʽ�Ƿ�����
    ' PrivateMode = "0" Ϊ�ر��޺�ģʽ��PrivateMode = "1" ����޺�ģʽ

    If PrivateMode = "0" Then
    
        Label1.Caption = "�޺�ģʽ - NetExplore ��ҳ�����"
        GotoNoHistory.Caption = "�ر��޺����(&H)"
        PrivateMode = "1"
        
    Else
    
        Label1.Caption = "ʹ�� Trident ��Ⱦ���� - NetExplore " & NEV
        PrivateMode = "0"
        GotoNoHistory.Caption = "�����޺����(&H)"
        
    End If
    
End Sub

Private Sub GotoSmallWindow_Click()             '����˵����е�С��ģʽ

    On Error Resume Next
    LitModForm.Visible = True
    Me.Visible = False
    LitModForm.Show

End Sub

Private Sub HistoryForm1_Click()                '����˵����е���ʷ��¼

    HistoryForm.Show 1

End Sub

Private Sub NewNE_Click()
    NewUpdForm.Show 1
End Sub

Private Sub SearchThing_Click()             '����˵����е�������ҳ��ť

    WebBrowser1.ExecWB OLECMDID_FIND, OLECMDEXECOPT_DODEFAULT

End Sub

Private Sub Settings_Click()                '����˵����е�����ѡ��

    ControlForm.Show 1                '�������

End Sub

Private Sub Text1_Change()

    If Not PrivateMode = "1" Then                   '���û�п����޺�ģʽ����
    
        Open App.Path & "\History\URL.dat" For Append As #2
        Print #2, WebBrowser1.LocationURL
        Close #2
        
        Open App.Path & "\History\time.dat" For Append As #2
        Print #2, Date & " " & Time
        Close #2
        
        Open App.Path & "\History\title.dat" For Append As #2
        Print #2, WebBrowser1.Document.Title
        Close #2
    
    End If
    
    'NE�Ĵ��ڱ�����ʾΪ��ҳ�ı���
    MainForm.Caption = " NetExplore " & NEV & " - " & WebBrowser1.Document.Title

End Sub

Private Sub Text1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)              '������ƶ���Text1�ı���ʱ

    If Text1.ForeColor = &H80000006 Then            '����ı�����ɫΪ��ɫ
        Text1.ForeColor = vbBlack                               '�ı�����ɫ��Ϊ��ɫ
    End If

End Sub

Private Sub Timer1_Timer()
NewUpdForm.Show 1
Timer1.Enabled = False
End Sub

Private Sub Command6_Click()            '�����ѡ���ťʱ

    PopupMenu control, , Command6.Left - 3400, Command6.Top + Command6.Height + 120          'ѡ�ť�����˵�

End Sub

Private Sub Command2_Click()                '�����ˢ�¡���ťʱ

    On Error Resume Next                '����� On Error Resume Next ��Ϊ�˷�ֹҳ��ĳЩ������޷�ˢ��ʱ��������ʵ�����������Բ���û���κ�Ӱ�죬ֱ�Ӻ��Ծͺ��ˣ�����Ҳ�����������ĵط���ͳһ�� On Error Resume Next
    WebBrowser1.Refresh

End Sub

Private Sub Command3_Click()                '��������ء���ťʱ

    On Error Resume Next
    WebBrowser1.GoBack
    
End Sub
Private Sub Command4_Click()                '�����ǰ������ťʱ

    On Error Resume Next
    WebBrowser1.GoForward

End Sub

Private Sub Command8_Click()                '�����������ҳ����ťʱ

    On Error Resume Next
    WebBrowser1.Navigate ReadInf(App.Path & "\neset.ini", "Home")

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

'�Զ����հ�

If Text3.Text = "" Then
    Text3.Text = " ʹ�� Microsoft Bing ����"
    Text3.ForeColor = &H808080
End If

If Text1.ForeColor = vbBlack Then
    Text1.ForeColor = &H80000006
End If

End Sub

Private Sub Form_Resize()
On Error Resume Next

'��������������ִ�С�Զ�����

WebBrowser1.Width = MainForm.Width - 200
WebBrowser1.Height = MainForm.Height - WebBrowser1.Top - 580
Text1.Width = (MainForm.Width - Text1.Left - 480) / 3 * 2
Text3.Left = Text1.Width + Text1.Left + 120
Text3.Width = (MainForm.Width - Text1.Left - 480) / 3

'����������ܻ���-����

Line1.X2 = MainForm.Width

'���������水ť����

Command6.Left = MainForm.Width - Command6.Width - 360
Command1.Left = Command6.Left - Command1.Width - 120
Command5.Left = Command1.Left - Command5.Width - 120
Command2.Left = Command5.Left - Command2.Width - 120
Command8.Left = Command2.Left - Command8.Width - 120

Label1.Width = Command8.Left - 240 - Label1.Left

Select Case Me.WindowState
Case 0
On Error Resume Next
Case 1
On Error Resume Next
Case 2
On Error Resume Next
End Select

End Sub
Private Sub Form_Load()

NEV = "NeXT 2"
Label1.Caption = "ʹ�� Trident ��Ⱦ���� - NetExplore " & NEV
MainForm.Caption = " NetExplore " & NEV

MainForm.Width = 15500                    '�����������С
MainForm.Height = 10745

Dim OpenNU As Boolean

OpenNU = False

If Dir(App.Path & "\neset.ini") = "" Then

    Open App.Path & "\neset.ini" For Append As #10
    Close #10
    
    Open App.Path & "\neset.ini" For Output As #11
    Print #11, "NEV=" & NEV
    Print #11, "Home=https://cn.bing.com/"
    Print #11, "Trident=11000"
    Close #11

    OpenNU = True

End If

If ReadInf(App.Path & "\neset.ini", "NEV") = "" Then
    Open App.Path & "\neset.ini" For Append As #11
    Print #11, "NEV=" & NEV
    Close #11
    OpenNU = True
End If

If Not ReadInf(App.Path & "\neset.ini", "NEV") = NEV Then
    PrintInf App.Path & "\neset.ini", "NEV", NEV
    OpenNU = True
End If

If OpenNU = True Then Timer1.Enabled = True

On Error Resume Next
MkDir App.Path & "\History"

PrivateMode = "0"               '�����޺�ģʽΪ�ر�

' ================= IE�ں����� =================

Dim w
    Set w = CreateObject("wscript.shell")
    w.regwrite "HKEY_CURRENT_USER\Software\Microsoft\Internet Explorer\Main\FeatureControl\FEATURE_BROWSER_EMULATION\" & App.EXEName + ".exe", ReadInf(App.Path & "\neset.ini", "Trident"), "REG_DWORD"
Set w = Nothing

    WebBrowser1.Navigate ReadInf(App.Path & "\neset.ini", "Home")

Dim webnet As String
webnet = VBA.Command

If Not webnet = "" Then
    WebBrowser1.Navigate webnet
End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub forward_Click()
On Error Resume Next
WebBrowser1.GoForward
End Sub

Private Sub fresh_Click()
On Error Resume Next
WebBrowser1.Refresh
End Sub


Private Sub home_Click()
On Error Resume Next
WebBrowser1.GoHome
End Sub

Private Sub F12open_Click()

Shell (VBA.Environ("systemroot") & "\system32\f12\IEChooser.exe"), 1

End Sub

Private Sub print_Click()
WebBrowser1.ExecWB OLECMDID_PRINT, OLECMDEXECOPT_DONTPROMPTUSER
End Sub

Private Sub QuanScreen_Click()
    Me.WindowState = 2
End Sub

Private Sub SF100_Click()
WebBrowser1.Document.body.Style.Zoom = "100%"
End Sub

Private Sub SF150_Click()
WebBrowser1.Document.body.Style.Zoom = "150%"
End Sub

Private Sub SF20_Click()
WebBrowser1.Document.body.Style.Zoom = "20%"
End Sub

Private Sub SF50_Click()
WebBrowser1.Document.body.Style.Zoom = "50%"
End Sub

Private Sub SF80_Click()
WebBrowser1.Document.body.Style.Zoom = "80%"
End Sub

Private Sub SFZDY_Click()
ZoomForm.Show 1
End Sub

Private Sub stop_Click()
On Error Resume Next
WebBrowser1.stop
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
On Error Resume Next

'���ûس�������

 If KeyAscii = 13 Then
        WebBrowser1.Navigate Text1.Text
    End If
    
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)

'������»س������������ҳ

If KeyAscii = 13 Then
    WebBrowser1.Navigate "https://cn.bing.com/search?q=" & (UTF8EncodeURI(Text3.Text))
End If
If Text3.Text = " ʹ�� Microsoft Bing ����" Then
Text3.ForeColor = &H0&
Text3.Text = ""
End If

End Sub
Function UTF8EncodeURI(szInput)                 'תUTF8������
Dim wch, uch, szRet
Dim X
Dim nAsc, nAsc2, nAsc3
If szInput = "" Then
UTF8EncodeURI = szInput
Exit Function
End If
For X = 1 To Len(szInput)
wch = Mid(szInput, X, 1)
nAsc = AscW(wch)
If nAsc < 0 Then nAsc = nAsc + 65536
If (nAsc And &HFF80) = 0 Then
szRet = szRet & wch
Else
If (nAsc And &HF000) = 0 Then
uch = "%" & Hex(((nAsc \ 2 ^ 6)) Or &HC0) & Hex(nAsc And &H3F Or &H80)
szRet = szRet & uch
Else
uch = "%" & Hex((nAsc \ 2 ^ 12) Or &HE0) & "%" & _
Hex((nAsc \ 2 ^ 6) And &H3F Or &H80) & "%" & _
Hex(nAsc And &H3F Or &H80)
szRet = szRet & uch
End If
End If
Next
UTF8EncodeURI = szRet
End Function

Function GBKEncodeURI(szInput)          '����GBK15����
Dim i As Long
Dim X() As Byte
Dim szRet As String
szRet = ""
X = StrConv(szInput, vbFromUnicode)
For i = LBound(X) To UBound(X)
szRet = szRet & "%" & Hex(X(i))
Next
GBKEncodeURI = szRet
End Function

Private Sub Text3_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next

'�������������������������ʧ

If Text3.Text = " ʹ�� Microsoft Bing ����" Then
Text3.ForeColor = &H0&
Text3.Text = ""
End If

End Sub

Private Sub WebBrowser1_GotFocus()

'�Զ����հ�

If Text3.Text = "" Then
    Text3.Text = " ʹ�� Microsoft Bing ����"
    Text3.ForeColor = &H808080
End If

If Text1.ForeColor = vbBlack Then
    Text1.ForeColor = &H80000006
End If

End Sub

Private Sub WebBrowser1_StatusTextChange(ByVal Text As String)
On Error Resume Next
End Sub
Private Sub WebBrowser1_DownloadBegin()
On Error Resume Next
WebBrowser1.Silent = True
End Sub

Private Sub WebBrowser1_DownloadComplete()
On Error Resume Next
WebBrowser1.Silent = True
End Sub
Private Sub WebBrowser1_NewWindow2(ppDisp As Object, Cancel As Boolean)
On Error Resume Next
Cancel = True
WebBrowser1.Navigate2 WebBrowser1.Document.activeElement.href
End Sub


Private Sub WebBrowser1_TitleChange(ByVal Text As String)
On Error Resume Next

'����������ҳ��ִ��

Text1.Text = " " & WebBrowser1.LocationURL

    If InternetGetConnectedState(0&, 0&) Then
    WebBrowser1.Visible = True
    Else
    WebBrowser1.Visible = False
    End If

If InStr(Text1.Text, "http://") <> 0 Then
Text2.Text = "����ȫ"
End If

If InStr(Text1.Text, "https://") <> 0 Then
Text2.Text = "��ȫ"
End If

If InStr(Text1.Text, "www.aroton.top") <> 0 Then
Shell "explorer.exe" & Text1.Text
WebBrowser1.GoHome
End If

End Sub
Private Sub Text1_GotFocus()
Text1.SelStart = 0
Text1.SelLength = Len(Text1)
End Sub

