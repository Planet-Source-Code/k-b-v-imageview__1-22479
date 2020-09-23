VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.DLL"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   Caption         =   "Image Viewer"
   ClientHeight    =   7185
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7665
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7185
   ScaleWidth      =   7665
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Left            =   6840
      Top             =   1920
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   6840
      Top             =   3000
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0896
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0CEA
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":113E
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":145A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":18AE
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1D02
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.FileListBox File1 
      Enabled         =   0   'False
      Height          =   2625
      Left            =   4800
      Pattern         =   "*.jpg;*.gif;*.bmp"
      TabIndex        =   2
      Top             =   4200
      Visible         =   0   'False
      Width           =   2175
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   7665
      _ExtentX        =   13520
      _ExtentY        =   1005
      ButtonWidth     =   1535
      ButtonHeight    =   953
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   8
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Browse"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Previous"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Refresh"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Next"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Slide show"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "About"
            ImageIndex      =   6
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin SHDocVwCtl.WebBrowser picView 
      Height          =   6495
      Left            =   0
      TabIndex        =   0
      Top             =   600
      Width           =   7575
      ExtentX         =   13361
      ExtentY         =   11456
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
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim stat As Integer
Dim strResFolder As String
Private Sub Form_Load()
stat = 0
picView.Navigate "about:blank"
Open App.Path & "\slideshow.tmp" For Output As #1
Print #1, ""
Close #1
Open App.Path & "\tn.tmp" For Output As #1
Print #1, ""
Close #1
End Sub

Private Sub Form_Resize()
picView.Height = Form1.ScaleHeight - picView.Top
picView.Width = Form1.ScaleWidth
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
Kill App.Path & "\slideshow.tmp"
Kill App.Path & "\tn.tmp"
End Sub

Private Sub Timer1_Timer()
Dim slideshow As String
If File1.ListIndex = File1.ListCount - 1 Then
Timer1.Enabled = False
stat = 0
picView.Navigate2 App.Path & "\tn.tmp"
Else
File1.Selected(stat) = True
slideshow = "<center><img src=""" & File1.Path & "\" & File1.FileName & """ border=""0""></center>"
Open App.Path & "\slideshow.tmp" For Output As #1
Print #1, slideshow
Close #1
picView.Navigate App.Path & "\slideshow.tmp"
stat = stat + 1
End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

Select Case Button.Caption
Case "Browse"
Dim htmlcode, finalcode, thumbnail As String
strResFolder = BrowseForFolder(hWnd, "Choose folder.")
If strResFolder = "" Then
    Exit Sub
Else
    File1.Path = strResFolder
    Form1.Caption = "Image Viewer - " & strResFolder
End If
For i = 0 To File1.ListCount - 1
File1.Selected(i) = True
htmlcode = "<a href=""" & File1.Path & "\" & File1.FileName & """><img src=""" & File1.Path & "\" & File1.FileName & """ width=""100"" height=""100"" border=""0""><BR></a> " & vbCrLf
finalcode = finalcode & htmlcode
Next i
thumbnail = "<HTML><BODY><CENTER> <BR>" & finalcode & "</center></BODY></html>"
Open App.Path & "\tn.tmp" For Output As #1
Print #1, thumbnail
Close #1
picView.Navigate2 App.Path & "\tn.tmp"
Case "Previous"
On Error Resume Next
picView.GoBack
Case "Refresh"
picView.Refresh
Case "Next"
On Error Resume Next
picView.GoForward
Toolbar1.Buttons.Item(3).Enabled = True
Case "Slide show"
If strResFolder <> "" Then
File1.Selected(0) = True
a = InputBox("Enter time interval in seconds:", "Time Interval")
If a = "" Then
Timer1.Interval = 5 * 1000
Timer1.Enabled = True
Else
Timer1.Interval = a * 1000
Timer1.Enabled = True
End If
Else
MsgBox "Choose a directory."
End If
Case "About"
MsgBox "Cool little image viewer." & vbCrLf & _
"This app has afew options. Has the ability to view animated gifs," & vbCrLf & _
"thumbnails, and slideshows. Of course, cuz its made from the" & vbCrLf & _
"IE Controls..ehe! And I just realized that you can set the selected" & vbCrLf & _
"picture as the wallpaper too! LoL! This app even have built in options..:)"
End Select
End Sub
