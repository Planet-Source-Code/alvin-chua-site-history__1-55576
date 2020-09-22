VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form Form3 
   Caption         =   "Browser"
   ClientHeight    =   5130
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6270
   Icon            =   "Browser.frx":0000
   LinkTopic       =   "Form3"
   ScaleHeight     =   5130
   ScaleWidth      =   6270
   StartUpPosition =   3  'Windows Default
   Begin SHDocVwCtl.WebBrowser wb 
      Height          =   5175
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6255
      ExtentX         =   11033
      ExtentY         =   9128
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
      Location        =   ""
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim OrHeight As Integer
Dim OrWidth As Integer
Private Sub Form_Load()
Me.Caption = "Loading..."
wb.Height = Me.Height - 500
wb.Width = Me.Width - 120
OrHeight = Me.Height
OrWidth = Me.Width
End Sub

Private Sub Form_Resize()
On Error Resume Next
If Me.Height <> OrHeight Then
wb.Height = Me.Height - 500
OrHeight = Me.Height

End If
If Me.Width <> OrWidth Then
wb.Width = Me.Width - 120
OrWidth = Me.Width

End If

End Sub

Private Sub wb_NewWindow2(ppDisp As Object, Cancel As Boolean)
Dim browser2 As New Form3
browser2.Show
Set ppDisp = browser2.wb.object
wb.RegisterAsBrowser = True
wb.Silent = True
End Sub

Private Sub wb_TitleChange(ByVal Text As String)
Me.Caption = Text
End Sub
