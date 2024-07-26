VERSION 5.00
Begin VB.Form HistoryForm 
   BackColor       =   &H8000000E&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "NeXT 历史记录"
   ClientHeight    =   7350
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5760
   LinkTopic       =   "Form10"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7350
   ScaleWidth      =   5760
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton Command3 
      Caption         =   "清除选中"
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
      Left            =   3480
      TabIndex        =   8
      Top             =   1440
      Width           =   1575
   End
   Begin VB.ListBox toaList 
      Height          =   1680
      Left            =   8520
      TabIndex        =   6
      Top             =   2760
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.ListBox TitleList 
      Height          =   1680
      Left            =   8520
      TabIndex        =   5
      Top             =   720
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.ListBox URLList 
      Height          =   1680
      Left            =   5760
      TabIndex        =   4
      Top             =   2760
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.ListBox TimeList 
      Height          =   1680
      Left            =   5760
      TabIndex        =   3
      Top             =   720
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.CommandButton Command2 
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
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   1440
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "全部清除"
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
      Left            =   1800
      TabIndex        =   1
      Top             =   1440
      Width           =   1575
   End
   Begin VB.ListBox MainList 
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5160
      Left            =   120
      TabIndex        =   0
      Top             =   2040
      Width           =   5535
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      BackStyle       =   0  'Transparent
      Caption         =   "NeXT 历史记录"
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
      TabIndex        =   7
      Top             =   360
      Width           =   3240
   End
   Begin VB.Image Image1 
      Height          =   1320
      Left            =   0
      Picture         =   "HistoryForm.frx":0000
      Top             =   0
      Width           =   6615
   End
End
Attribute VB_Name = "HistoryForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Function History_Load()

MainList.Clear

'刷新历史记录 URL
    
        Dim URLs
        URLList.Clear
        Open App.Path & "\History\URL.dat" For Append As #8
        Close #8
        
        Open App.Path & "\History\URL.dat" For Input As #1
            Do While Not EOF(1)
                Line Input #1, URLs
                URLList.AddItem URLs
            Loop
        Close #1
        
            '倒置 URLs List 框
            
        For i = URLList.ListCount - 1 To 0 Step -1
            toaList.AddItem URLList.List(i)
        Next
        
        URLList.Clear
        
        For i = 0 To toaList.ListCount
            URLList.AddItem toaList.List(i)
        Next
        
        toaList.Clear

'刷新历史记录时间 time
    
        Dim HisTime
        TimeList.Clear
        Open App.Path & "\History\time.dat" For Append As #8
        Close #8
        
        Open App.Path & "\History\time.dat" For Input As #1
            Do While Not EOF(1)
                Line Input #1, HisTime
                TimeList.AddItem HisTime
            Loop
        Close #1
    
            '倒置 Time List 框
            
        For i = TimeList.ListCount - 1 To 0 Step -1
            toaList.AddItem TimeList.List(i)
        Next
        
        TimeList.Clear
        
        For i = 0 To toaList.ListCount
            TimeList.AddItem toaList.List(i)
        Next
        
        toaList.Clear

'刷新历史记录标题 title

        Dim HisTitle
        TitleList.Clear
        Open App.Path & "\History\title.dat" For Append As #8
        Close #8
        
        Open App.Path & "\History\title.dat" For Input As #1
            Do While Not EOF(1)
                Line Input #1, HisTitle
                TitleList.AddItem HisTitle
            Loop
        Close #1
        
            '倒置 Title List 框
        
        For i = TitleList.ListCount - 1 To 0 Step -1
            toaList.AddItem TitleList.List(i)
        Next
        
        TitleList.Clear
        
        For i = 0 To toaList.ListCount
            TitleList.AddItem toaList.List(i)
        Next
        
        toaList.Clear

'填充mainList

Dim hstMain
Dim hstTA
Dim hstTB

For i = 0 To TimeList.ListCount
    hstTA = Left(TimeList.List(i), 8)
    hstTB = Left(Right(TimeList.List(i), 8), 5)
    hstMain = hstTA & " " & hstTB & " | " & TitleList.List(i)
    If Not (hstMain = "  | ") Then
    MainList.AddItem hstMain
    End If
Next

End Function

Private Sub Command1_Click()

Dim okcam As String
okcam = MsgBox("此操作不可挽回！确定要清除全部历史记录吗？", vbOKCancel + vbInformation, "确认当前操作")

    If okcam = vbOK Then
    
        Kill App.Path & "\History\URL.dat"
        Open App.Path & "\History\URL.dat" For Append As #2
        Close #2
        
        Kill App.Path & "\History\time.dat"
        Open App.Path & "\History\time.dat" For Append As #2
        Close #2
        
        Kill App.Path & "\History\title.dat"
        Open App.Path & "\History\title.dat" For Append As #2
        Close #2

    '重新加载历史记录
    History_Load
        
    End If
    
End Sub

Private Sub Command2_Click()

'重新加载历史记录
History_Load

End Sub

Private Sub Command3_Click()
    On Error Resume Next
    TimeList.RemoveItem MainList.ListIndex
    URLList.RemoveItem MainList.ListIndex
    TitleList.RemoveItem MainList.ListIndex
    MainList.RemoveItem MainList.ListIndex
    
    Open App.Path & "\History\URL.dat" For Output As #25
    For i = 0 To URLList.ListCount - 1
    Print #25, URLList.List(i)
    Next i
    Close #25
    
    Open App.Path & "\History\time.dat" For Output As #25
    For i = 0 To TimeList.ListCount - 1
    Print #25, TimeList.List(i)
    Next i
    Close #25
        
    Open App.Path & "\History\title.dat" For Output As #25
    For i = 0 To TitleList.ListCount - 1
    Print #25, TitleList.List(i)
    Next i
    Close #25

    History_Load

End Sub

Private Sub Form_Load()
Me.Icon = LoadPicture("")

'加载历史记录
History_Load
    
End Sub

Private Sub MainList_DblClick()
HistOpenForm.Show 1
End Sub
