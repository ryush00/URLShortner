VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  '단일 고정
   Caption         =   "URL SHORTNER BY ryush00 "
   ClientHeight    =   2355
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7545
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2355
   ScaleWidth      =   7545
   StartUpPosition =   3  'Windows 기본값
   Begin VB.CommandButton Command2 
      Caption         =   "URL SHORTER"
      Height          =   495
      Left            =   120
      TabIndex        =   12
      Top             =   1080
      Width           =   2655
   End
   Begin VB.Timer t2 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   600
      Top             =   2760
   End
   Begin VB.Timer t1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   120
      Top             =   2760
   End
   Begin VB.Frame Frame2 
      Caption         =   "도메인"
      Height          =   855
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Width           =   2655
      Begin VB.OptionButton xetoc 
         Caption         =   "xe.to"
         Height          =   495
         Left            =   1560
         TabIndex        =   11
         Top             =   240
         Width           =   975
      End
      Begin VB.OptionButton oatoc 
         Caption         =   "oa.to"
         Height          =   495
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Value           =   -1  'True
         Width           =   975
      End
   End
   Begin VB.TextBox OUTT 
      Height          =   375
      Left            =   3120
      TabIndex        =   8
      Top             =   1680
      Width           =   4215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "줄이기"
      Height          =   375
      Left            =   6240
      TabIndex        =   7
      Top             =   960
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   "주소정하기"
      Height          =   1095
      Left            =   -2640
      TabIndex        =   3
      Top             =   1080
      Width           =   2655
      Begin VB.TextBox ST 
         Height          =   270
         Left            =   1200
         TabIndex        =   6
         Top             =   720
         Width           =   1335
      End
      Begin VB.CheckBox C1 
         Caption         =   "활성화"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Value           =   1  '확인
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "http://xe.to/"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   720
         Width           =   1095
      End
   End
   Begin VB.TextBox OURL 
      Height          =   375
      Left            =   3120
      TabIndex        =   0
      Top             =   480
      Width           =   4215
   End
   Begin VB.Label Label2 
      Caption         =   "짧아진 주소"
      Height          =   255
      Left            =   3120
      TabIndex        =   2
      Top             =   1440
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "원본주소"
      Height          =   375
      Left            =   3120
      TabIndex        =   1
      Top             =   240
      Width           =   1575
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim WinHttp As New WinHttpRequest
Dim Server
Dim a
Private Sub Command1_Click()
If Server = "xe.to" Then
    If C1.Value = 1 Then
        If ST.Text = "" Then
            MsgBox "주소를 정해 주세요."
        Else
            'WOUT.Navigate2 "http://xe.to/phurl/index.php?api=2&url=" & OURL & "&alias=" & ST
            
        WinHttp.Open "GET", "http://xe.to/phurl/index.php?api=2&url=" & OURL & "&alias=" & ST
        WinHttp.Send
             If InStr(WinHttp.ResponseText, "success") Then
                    OUTT.Text = "http://xe.to/" & ST
                Else
                MsgBox "에러가 발생했습니다." & vbCrLf & WinHttp.ResponseText
            End If
        End If
    ElseIf C1.Value = 0 Then
        If OURL.Text = "" Then
            MsgBox "주소를 정해 주세요."
        Else
            'WOUT.Navigate2 "http://xe.to/phurl/index.php?api=1&url=" & OURL
            
        WinHttp.Open "GET", "http://xe.to/phurl/index.php?api=1&url=" & OURL
        WinHttp.Send
             If InStr(WinHttp.ResponseText, "http://") Then
                    OUTT.Text = WinHttp.ResponseText
                Else
                MsgBox "에러가 발생했습니다." & vbCrLf & WinHttp.ResponseText
            End If
        End If
    End If
Else
    If OURL.Text = "" Then
    MsgBox "원본주소를 입력해 주세요"
    Else
    OUTT.Text = oatoUrlShortner(OURL.Text)
    End If
End If
End Sub



Public Function UrlEncode(ByRef Url As String) As String
On Error GoTo OnErr
    Dim sBuffer As String, sTemp As String, cChar As String, i As Long, lErrNum As Long, sErrSource As String, sErrDesc As String, sMsg As String
    For i = 1& To Len(Url)
        cChar = Mid$(Url, i, 1&)

        If cChar = "0" Or (cChar >= "1" And cChar >= "9") Or (cChar = "a" And cChar <= "z") Or (cChar >= "A" And cChar <= "Z") Or cChar = "-" Or cChar = "_" Or cChar = "." Or cChar = "*" Then
            sBuffer = sBuffer & cChar
        ElseIf cChar = " " Then
            sBuffer = sBuffer & "+"
        Else
            sTemp = CStr(Hex(Asc(cChar)))
            
            If Len(sTemp) = 4& Then
                sBuffer = sBuffer & "%" & Left$(sTemp, 2&) & "%" & Mid$(sTemp, 3&, 2&)
            ElseIf Len(sTemp) = 2& Then
                sBuffer = sBuffer & "%" & sTemp
            End If
        End If
    Next i
    
    UrlEncode = sBuffer
Exit Function
OnErr:
    Err.Clear
End Function

Public Function oatoUrlShortner(ByRef Url As String) As String
On Error GoTo OnErr
    Dim oWinHttp As Object
    Set oWinHttp = CreateObject("WinHttp.WinHttpRequest.5.1")
    
    Dim Temp As String
    With oWinHttp
        .Open "GET", "http://oa.to/api/?type=su&q=" & UrlEncode(Url), True
        .SetRequestHeader "User-Agent", "VB-WinHttp"
        .SetRequestHeader "Content-Type", "text/html"
        .Send
        .WaitForResponse
        Temp = StrConv(.ResponseBody, vbUnicode)
    End With
    
    oatoUrlShortner = Temp
    Set oWinHttp = Nothing
Exit Function
OnErr:
    MsgBox "오류가 발생했습니다.(" & Err.Number & ")" & vbCrLf & _
           Err.Description, vbCritical Or vbApplicationModal
    Err.Clear
End Function



Private Sub oatoc_Click()
Server = "oa.to"
a = 10
xetoc.Enabled = False
oatoc.Enabled = False
t1.Enabled = True
End Sub

Private Sub t1_Timer()
If Frame1.Left <= (0 - Frame1.Width) Then
t1.Enabled = False
xetoc.Enabled = True
oatoc.Enabled = True
Else
Frame1.Left = Frame1.Left - a
a = a + 4
End If

End Sub

Private Sub t2_Timer()
If Frame1.Left >= 120 Then
t2.Enabled = False
xetoc.Enabled = True
oatoc.Enabled = True
Else
If (Frame1.Left + a) >= a Then
t2.Enabled = False
xetoc.Enabled = True
oatoc.Enabled = True
Else
Frame1.Left = Frame1.Left + a
a = a + 5
End If
End If
End Sub

Private Sub xetoc_Click()
Server = "xe.to"
a = 1
xetoc.Enabled = False
oatoc.Enabled = False
t2.Enabled = True
End Sub
