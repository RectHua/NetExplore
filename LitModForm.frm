VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form LitModForm 
   BackColor       =   &H8000000E&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   7335
   ClientLeft      =   16980
   ClientTop       =   8490
   ClientWidth     =   11055
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7335
   ScaleWidth      =   11055
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'ÆÁÄ»ÖÐÐÄ
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   6615
      Left            =   120
      TabIndex        =   9
      Top             =   600
      Width           =   10815
      ExtentX         =   19076
      ExtentY         =   11668
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
   Begin VB.CommandButton Command5 
      Caption         =   "¹ØÓÚ"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8880
      TabIndex        =   8
      ToolTipText     =   "¹ØÓÚä¯ÀÀÆ÷"
      Top             =   120
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1320
      TabIndex        =   7
      Top             =   120
      Width           =   4575
   End
   Begin VB.CommandButton Command7 
      Caption         =   "ÍË³öÐ¡´°Ä£Ê½"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9600
      TabIndex        =   6
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Í£Ö¹"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7440
      TabIndex        =   5
      ToolTipText     =   "Í£Ö¹Ò³Ãæ¼ÓÔØ"
      Top             =   120
      Width           =   615
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Ö÷Ò³"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6000
      TabIndex        =   4
      ToolTipText     =   "·µ»Øä¯ÀÀÆ÷Ö÷Ò³"
      Top             =   120
      Width           =   615
   End
   Begin VB.CommandButton Command4 
      Caption         =   "¡ú"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   720
      TabIndex        =   3
      ToolTipText     =   "Ç°½ø"
      Top             =   120
      Width           =   495
   End
   Begin VB.CommandButton Command3 
      Caption         =   "¡û"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   2
      ToolTipText     =   "ºóÍË"
      Top             =   120
      Width           =   495
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Ë¢ÐÂ"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6720
      TabIndex        =   1
      ToolTipText     =   "ÖØÐÂ¼ÓÔØÒ³Ãæ"
      Top             =   120
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "±£´æ"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8160
      TabIndex        =   0
      ToolTipText     =   "Áí´æÎªÍøÒ³"
      Top             =   120
      Width           =   615
   End
   Begin VB.Shape Shape2 
      Height          =   6645
      Left            =   100
      Top             =   580
      Width           =   10845
   End
End
Attribute VB_Name = "LitModForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function InternetGetConnectedState Lib "wininet.dll" (ByRef dwFlags As Long, ByVal dwReserved As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Const HWND_TOPMOST& = -1
Private Const SWP_NOSIZE& = &H1
Private Const SWP_NOMOVE& = &H2

Private Sub Command1_Click()
WebBrowser1.ExecWB OLECMDID_SAVEAS, OLECMDEXECOPT_DODEFAULT
End Sub

Private Sub Command2_Click()
On Error Resume Next
WebBrowser1.Refresh
End Sub

Private Sub Command3_Click()
On Error Resume Next
  WebBrowser1.GoBack
End Sub

Private Sub Command4_Click()
On Error Resume Next
WebBrowser1.GoForward
End Sub

Private Sub Command5_Click()
AboutForm.Show
End Sub

Private Sub Command6_Click()
On Error Resume Next
WebBrowser1.Navigate homes
End Sub

Private Sub Command7_Click()
On Error Resume Next
MainForm.Visible = True
Me.Visible = False
MainForm.WebBrowser1.Navigate LitModForm.Text1.Text
End Sub

Private Sub Command8_Click()
On Error Resume Next
WebBrowser1.stop

End Sub



Private Sub Form_Load()
On Error Resume Next
WebBrowser1.Navigate MainForm.Text1.Text
Me.Icon = LoadPicture("")
Me.Caption = " NetExplore " & NEV
'SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
End Sub
Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
MainForm.Visible = True
Me.Visible = False
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
On Error Resume Next
 If KeyAscii = 13 Then
        WebBrowser1.Navigate Text1.Text
    End If
End Sub


Private Sub WebBrowser1_StatusTextChange(ByVal Text As String)
On Error Resume Next
End Sub
Private Sub WebBrowser1_DownloadBegin()
WebBrowser1.Silent = True
End Sub

Private Sub WebBrowser1_DownloadComplete()
WebBrowser1.Silent = True
End Sub
Private Sub WebBrowser1_NewWindow2(ppDisp As Object, Cancel As Boolean)
On Error Resume Next
Cancel = True
WebBrowser1.Navigate2 WebBrowser1.Document.activeElement.href
End Sub


Private Sub WebBrowser1_TitleChange(ByVal Text As String)
Text1.Text = WebBrowser1.LocationURL
End Sub
