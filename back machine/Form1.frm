VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Search The  History"
   ClientHeight    =   4785
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6240
   ForeColor       =   &H00FFFFFF&
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4785
   ScaleWidth      =   6240
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List1 
      Height          =   2985
      Left            =   240
      TabIndex        =   2
      Top             =   1200
      Width           =   5775
   End
   Begin VB.Frame Frame1 
      Height          =   375
      Left            =   0
      TabIndex        =   3
      Top             =   4440
      Width           =   6255
      Begin VB.CommandButton Command2 
         Caption         =   "About"
         Height          =   255
         Left            =   5160
         TabIndex        =   5
         Top             =   120
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "Done Loading"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   120
         Width           =   1455
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Go"
      Height          =   375
      Left            =   4920
      TabIndex        =   1
      Top             =   600
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   240
      TabIndex        =   0
      Top             =   720
      Width           =   4575
   End
   Begin VB.Image Image1 
      Height          =   420
      Left            =   240
      Picture         =   "Form1.frx":1B06
      Top             =   120
      Width           =   4020
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function InternetOpen Lib "wininet.dll" Alias "InternetOpenA" (ByVal sAgent As String, ByVal lAccessType As Long, ByVal sProxyName As String, ByVal sProxyBypass As String, ByVal lFlags As Long) As Long
Private Declare Function InternetOpenUrl Lib "wininet.dll" Alias "InternetOpenUrlA" (ByVal hInternetSession As Long, ByVal sURL As String, ByVal sHeaders As String, ByVal lHeadersLength As Long, ByVal lFlags As Long, ByVal lContext As Long) As Long
Private Declare Function InternetReadFile Lib "wininet.dll" (ByVal hFile As Long, ByVal sBuffer As String, ByVal lNumBytesToRead As Long, lNumberOfBytesRead As Long) As Integer
Private Declare Function InternetCloseHandle Lib "wininet.dll" (ByVal hInet As Long) As Integer
Private Const IF_FROM_CACHE = &H1000000
Private Const IF_MAKE_PERSISTENT = &H2000000
Private Const IF_NO_CACHE_WRITE = &H4000000
Private Const BUFFER_LEN = 256
Dim ar3() As String
Dim Data As Variant

Private Sub Command1_Click()
Label2.Caption = "Getting Data..."
DoEvents
Dim url As String
List1.Clear
Text1.Enabled = False
Command1.Enabled = False
url = Replace(Text1.Text, "http://", "")
url = Replace(url, "www.", "")
Data = GetUrlSource("http://web.archive.org/web/*/" & url)
Dim month As String
Dim ar() As String
Dim ar2() As String
Dim link As String
Dim mar() As String
link = ""
month = "jan feb mar apr may jun jul aug sep oct nov dec"
mar = Split(month, " ")
ar = Split(Data, "<a href=")
For i = LBound(ar) To UBound(ar)
ar2 = Split(ar(i), ">")
For k = LBound(ar2) To UBound(ar2)
For t = LBound(mar) To UBound(mar)
If InStr(1, LCase(ar2(k)), mar(t)) = 1 Then
link = link & ar2(k - 1) & vbCrLf
List1.AddItem Replace(ar2(k), "</a", "")
End If
Next t
Next k
Next i
ar3 = Split(link, vbCrLf)
If link = "" Then
MsgBox "No History Found"
Text1.Enabled = True
Command1.Enabled = True
Label2.Caption = "Done Loading"
End If
Text1.Enabled = True
Command1.Enabled = True
Label2.Caption = "Done Loading"

End Sub

Private Sub Command2_Click()
MsgBox "Program created by Peng, 13/3"
End Sub

Private Sub List1_DblClick()
Dim browser As New Form3
Load browser
browser.Show
browser.wb.Navigate Replace(ar3(List1.ListIndex), """", "")
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Call Command1_Click
End If
End Sub

Public Function GetUrlSource(sURL As String) As String
    Dim sBuffer As String * BUFFER_LEN, iResult As Integer, sData As String
    Dim hInternet As Long, hSession As Long, lReturn As Long
    hSession = InternetOpen("vb wininet", 1, vbNullString, vbNullString, 0)
    If hSession Then hInternet = InternetOpenUrl(hSession, sURL, vbNullString, 0, IF_NO_CACHE_WRITE, 0)
    If hInternet Then
        iResult = InternetReadFile(hInternet, sBuffer, BUFFER_LEN, lReturn)
        sData = sBuffer
        Do While lReturn <> 0
            iResult = InternetReadFile(hInternet, sBuffer, BUFFER_LEN, lReturn)
            sData = sData + Mid(sBuffer, 1, lReturn)
        Loop
    End If
    iResult = InternetCloseHandle(hInternet)
    GetUrlSource = sData
End Function

