VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5670
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   8460
   LinkTopic       =   "Form1"
   ScaleHeight     =   5670
   ScaleWidth      =   8460
   StartUpPosition =   3  '窗口缺省
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   4815
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   8175
      ExtentX         =   14420
      ExtentY         =   8493
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
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   5160
      TabIndex        =   0
      Top             =   5040
      Width           =   2295
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
TencentSDK.init ("AppKey") '这是个示例，请把AppKey换成您自己的AppKey
WebBrowser1.Navigate (TencentSDK.GetAuthPage("CallBackURL")) '请把CallBack换成您自己的CallBack
End Sub

Private Sub WebBrowser1_TitleChange(ByVal Text As String)
If InStr(WebBrowser1.LocationURL, "access_token") Then
TencentSDK.GetAccessToken (WebBrowser1.LocationURL)
Me.Hide
Form2.Show
Unload Me
End If
End Sub
