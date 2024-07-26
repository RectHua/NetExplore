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
   StartUpPosition =   2  '屏幕中心
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   8520
      Top             =   120
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "微软雅黑"
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
         Name            =   "微软雅黑"
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
      Text            =   "安全"
      Top             =   650
      Width           =   855
   End
   Begin VB.CommandButton Command5 
      Caption         =   "收藏"
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
      Left            =   14400
      TabIndex        =   17
      Top             =   120
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "历史"
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
         Name            =   "微软雅黑"
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
      Text            =   " 使用 Microsoft Bing 搜索"
      Top             =   650
      Width           =   5295
   End
   Begin VB.CommandButton Command4 
      Caption         =   "→"
      BeginProperty Font 
         Name            =   "微软雅黑"
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
      ToolTipText     =   "当遇到困苦的时候，应该迎难而上，而不是放弃自我。"
      Top             =   650
      Width           =   615
   End
   Begin VB.CommandButton Command3 
      Caption         =   "←"
      BeginProperty Font 
         Name            =   "微软雅黑"
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
      ToolTipText     =   "有时候，后退是最明确的选择。"
      Top             =   650
      Width           =   615
   End
   Begin VB.CommandButton Command8 
      Caption         =   "主页"
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
      Left            =   12720
      TabIndex        =   2
      Top             =   120
      Width           =   735
   End
   Begin VB.CommandButton Command6 
      Caption         =   "选项"
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
      Left            =   16080
      TabIndex        =   1
      Top             =   120
      Width           =   735
   End
   Begin VB.CommandButton Command2 
      Caption         =   "刷新"
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
      Left            =   13560
      TabIndex        =   0
      Top             =   120
      Width           =   735
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "连接到互联网，Internet 因为有你而不同。您可以："
      BeginProperty Font 
         Name            =   "微软雅黑 Light"
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
      Caption         =   "| 试着重启一下无线路由器，宽带或信号交换机。"
      BeginProperty Font 
         Name            =   "微软雅黑 Light"
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
      Caption         =   "| 看看有没有关闭飞行模式。"
      BeginProperty Font 
         Name            =   "微软雅黑 Light"
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
      Caption         =   "| 检查确认所有网线是否都插好了。"
      BeginProperty Font 
         Name            =   "微软雅黑 Light"
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
      Caption         =   "| 刷新当前页面或尝试重新访问。"
      BeginProperty Font 
         Name            =   "微软雅黑 Light"
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
      Caption         =   "| 确保 NetExplore 没有被杀毒软件和防火墙拦截。"
      BeginProperty Font 
         Name            =   "微软雅黑 Light"
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
      Caption         =   "| 确保路由器在正常工作并且信号稳定。"
      BeginProperty Font 
         Name            =   "微软雅黑 Light"
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
      Caption         =   "嗯... NetExplore 找不到 Internet"
      BeginProperty Font 
         Name            =   "微软雅黑 Light"
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
      ToolTipText     =   "致敬互联网时代先驱 - Internet Explorer"
      Top             =   105
      Width           =   960
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "使用 Trident 渲染引擎 - NetExplore 50"
      BeginProperty Font 
         Name            =   "微软雅黑 Light"
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
      Caption         =   "控制面板"
      Visible         =   0   'False
      Begin VB.Menu WYE 
         Caption         =   "网页(&H)"
         Begin VB.Menu BACK 
            Caption         =   "返回(&B)"
            Shortcut        =   ^B
         End
         Begin VB.Menu forward 
            Caption         =   "前进(&F)"
            Shortcut        =   ^N
         End
         Begin VB.Menu stop 
            Caption         =   "停止(&O)"
            Shortcut        =   ^O
         End
         Begin VB.Menu fresh 
            Caption         =   "刷新(&R)"
            Shortcut        =   ^K
         End
         Begin VB.Menu home 
            Caption         =   "主页(&H)"
            Shortcut        =   ^H
         End
      End
      Begin VB.Menu prin 
         Caption         =   "打印(&P)"
         Begin VB.Menu print 
            Caption         =   "打印(P)..."
            Shortcut        =   ^P
         End
      End
      Begin VB.Menu file 
         Caption         =   "文件(&F)"
         Begin VB.Menu QuanScreen 
            Caption         =   "全屏(&L)      "
            Shortcut        =   {F11}
         End
         Begin VB.Menu LCW 
            Caption         =   "另存为(&A)..."
            Shortcut        =   ^S
         End
         Begin VB.Menu SearchThing 
            Caption         =   "搜索网页内容(S)..."
         End
      End
      Begin VB.Menu Display 
         Caption         =   "显示(&D)"
         Begin VB.Menu ZoomSetUp 
            Caption         =   "放大 Ctrl +（或滚轮加）"
         End
         Begin VB.Menu SmallSuoxiao 
            Caption         =   "缩小 Ctrl -（或滚轮减）"
         End
         Begin VB.Menu ThereIsNotLine 
            Caption         =   "-"
         End
         Begin VB.Menu SF150 
            Caption         =   "缩放 150%"
         End
         Begin VB.Menu SF100 
            Caption         =   "缩放 100%"
         End
         Begin VB.Menu SF80 
            Caption         =   "缩放 80%"
         End
         Begin VB.Menu SF50 
            Caption         =   "缩放 50%"
         End
         Begin VB.Menu SF20 
            Caption         =   "缩放 20%"
         End
         Begin VB.Menu NoneLineMenu 
            Caption         =   "-"
         End
         Begin VB.Menu SFZDY 
            Caption         =   "自定义缩放"
            Shortcut        =   ^Z
         End
      End
      Begin VB.Menu Nones 
         Caption         =   "-"
         Index           =   0
      End
      Begin VB.Menu Settings 
         Caption         =   "设置(&S)"
      End
      Begin VB.Menu ModeSystem 
         Caption         =   "视图(&M)"
         Begin VB.Menu GotoSmallWindow 
            Caption         =   "小窗模式(&S)"
         End
         Begin VB.Menu GotoNoHistory 
            Caption         =   "开启无痕浏览(&H)"
         End
      End
      Begin VB.Menu TOOLS 
         Caption         =   "工具(&T)"
         Begin VB.Menu F12open 
            Caption         =   "F12开发者工具(&L)"
            Shortcut        =   ^Y
         End
         Begin VB.Menu FavoritePage 
            Caption         =   "网页收藏夹(&F)"
         End
         Begin VB.Menu HistoryForm1 
            Caption         =   "浏览历史(&H)"
            Shortcut        =   ^U
         End
      End
      Begin VB.Menu None 
         Caption         =   "-"
      End
      Begin VB.Menu NewNE 
         Caption         =   "NE 新增功能（&N）"
      End
      Begin VB.Menu about 
         Caption         =   "关于 NetWork Explore (&A)"
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

Private Sub ABOUT_Click()           '点击菜单栏中的关于NE

    AboutForm.Show 1
    '保持最前显示关于窗口
    
End Sub

Private Sub BACK_Click()            '点击菜单栏中的返回按钮

    On Error Resume Next
    WebBrowser1.GoBack
    
End Sub

Private Sub LCW_Click()            '点击菜单栏中的另存为选项

    WebBrowser1.ExecWB OLECMDID_SAVEAS, OLECMDEXECOPT_DODEFAULT
    
End Sub

Private Sub Command1_Click()            '点击“历史”按钮

    HistoryForm.Show 1

End Sub
Private Sub Command5_Click()                '点击“收藏”按钮

    FavoriteForm.Show 1

End Sub

Private Sub FavoritePage_Click()                '点击菜单栏中的“网页收藏夹”
    
    FavoriteForm.Show 1

End Sub

Private Sub GotoNoHistory_Click()               '点击菜单栏中的启动无痕模式

    ' 本块中有一个名为 PrivateMode 的 String 变量，已经在通用部分中声明，作用为记录无痕模式是否启动
    ' PrivateMode = "0" 为关闭无痕模式，PrivateMode = "1" 则打开无痕模式

    If PrivateMode = "0" Then
    
        Label1.Caption = "无痕模式 - NetExplore 网页浏览器"
        GotoNoHistory.Caption = "关闭无痕浏览(&H)"
        PrivateMode = "1"
        
    Else
    
        Label1.Caption = "使用 Trident 渲染引擎 - NetExplore " & NEV
        PrivateMode = "0"
        GotoNoHistory.Caption = "开启无痕浏览(&H)"
        
    End If
    
End Sub

Private Sub GotoSmallWindow_Click()             '点击菜单栏中的小窗模式

    On Error Resume Next
    LitModForm.Visible = True
    Me.Visible = False
    LitModForm.Show

End Sub

Private Sub HistoryForm1_Click()                '点击菜单栏中的历史记录

    HistoryForm.Show 1

End Sub

Private Sub NewNE_Click()
    NewUpdForm.Show 1
End Sub

Private Sub SearchThing_Click()             '点击菜单栏中的搜索网页按钮

    WebBrowser1.ExecWB OLECMDID_FIND, OLECMDEXECOPT_DODEFAULT

End Sub

Private Sub Settings_Click()                '点击菜单栏中的设置选项

    ControlForm.Show 1                '控制面板

End Sub

Private Sub Text1_Change()

    If Not PrivateMode = "1" Then                   '如果没有开启无痕模式，则
    
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
    
    'NE的窗口标题显示为网页的标题
    MainForm.Caption = " NetExplore " & NEV & " - " & WebBrowser1.Document.Title

End Sub

Private Sub Text1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)              '当鼠标移动过Text1文本框时

    If Text1.ForeColor = &H80000006 Then            '如果文本框颜色为灰色
        Text1.ForeColor = vbBlack                               '文本的颜色变为黑色
    End If

End Sub

Private Sub Timer1_Timer()
NewUpdForm.Show 1
Timer1.Enabled = False
End Sub

Private Sub Command6_Click()            '点击“选项”按钮时

    PopupMenu control, , Command6.Left - 3400, Command6.Top + Command6.Height + 120          '选项按钮弹出菜单

End Sub

Private Sub Command2_Click()                '点击“刷新”按钮时

    On Error Resume Next                '这里的 On Error Resume Next 是为了防止页面某些情况下无法刷新时弹出报错，实际上这个报错对操作没有任何影响，直接忽略就好了，下面也有类似这样的地方，统一用 On Error Resume Next
    WebBrowser1.Refresh

End Sub

Private Sub Command3_Click()                '点击“返回”按钮时

    On Error Resume Next
    WebBrowser1.GoBack
    
End Sub
Private Sub Command4_Click()                '点击“前进”按钮时

    On Error Resume Next
    WebBrowser1.GoForward

End Sub

Private Sub Command8_Click()                '点击“返回主页”按钮时

    On Error Resume Next
    WebBrowser1.Navigate ReadInf(App.Path & "\neset.ini", "Home")

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

'自动填充空白

If Text3.Text = "" Then
    Text3.Text = " 使用 Microsoft Bing 搜索"
    Text3.ForeColor = &H808080
End If

If Text1.ForeColor = vbBlack Then
    Text1.ForeColor = &H80000006
End If

End Sub

Private Sub Form_Resize()
On Error Resume Next

'设置浏览器各部分大小自动更改

WebBrowser1.Width = MainForm.Width - 200
WebBrowser1.Height = MainForm.Height - WebBrowser1.Top - 580
Text1.Width = (MainForm.Width - Text1.Left - 480) / 3 * 2
Text3.Left = Text1.Width + Text1.Left + 120
Text3.Width = (MainForm.Width - Text1.Left - 480) / 3

'在浏览器四周划线-美观

Line1.X2 = MainForm.Width

'设置主界面按钮靠齐

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
Label1.Caption = "使用 Trident 渲染引擎 - NetExplore " & NEV
MainForm.Caption = " NetExplore " & NEV

MainForm.Width = 15500                    '设置浏览器大小
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

PrivateMode = "0"               '设置无痕模式为关闭

' ================= IE内核设置 =================

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

'设置回车键功能

 If KeyAscii = 13 Then
        WebBrowser1.Navigate Text1.Text
    End If
    
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)

'如果按下回车键，则访问网页

If KeyAscii = 13 Then
    WebBrowser1.Navigate "https://cn.bing.com/search?q=" & (UTF8EncodeURI(Text3.Text))
End If
If Text3.Text = " 使用 Microsoft Bing 搜索" Then
Text3.ForeColor = &H0&
Text3.Text = ""
End If

End Sub
Function UTF8EncodeURI(szInput)                 '转UTF8码声明
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

Function GBKEncodeURI(szInput)          '声明GBK15编码
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

'如果单击搜索框，让搜索文字消失

If Text3.Text = " 使用 Microsoft Bing 搜索" Then
Text3.ForeColor = &H0&
Text3.Text = ""
End If

End Sub

Private Sub WebBrowser1_GotFocus()

'自动填充空白

If Text3.Text = "" Then
    Text3.Text = " 使用 Microsoft Bing 搜索"
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

'浏览器浏览新页面执行

Text1.Text = " " & WebBrowser1.LocationURL

    If InternetGetConnectedState(0&, 0&) Then
    WebBrowser1.Visible = True
    Else
    WebBrowser1.Visible = False
    End If

If InStr(Text1.Text, "http://") <> 0 Then
Text2.Text = "不安全"
End If

If InStr(Text1.Text, "https://") <> 0 Then
Text2.Text = "安全"
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

