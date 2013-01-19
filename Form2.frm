VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "接口测试"
   ClientHeight    =   3990
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   6840
   LinkTopic       =   "Form2"
   ScaleHeight     =   3990
   ScaleWidth      =   6840
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   855
      Left            =   3600
      TabIndex        =   1
      Top             =   2880
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      Height          =   2175
      Left            =   240
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   240
      Width           =   6255
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
Text1.Text = TencentSDK.CallAPIPic("t/add_pic", Array("format", "content"), Array("json", Text1.Text), App.Path & "\0.png", "pic", "image/png", "0.png") '只是个示例
End Sub
