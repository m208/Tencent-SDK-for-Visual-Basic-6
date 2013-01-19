VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Tapi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'************ Tencent Weibo SDK for Visual Basic 6 ( OAuth 2 ) ************
'************                                                  ************
'************ 此 SDK 由 m208 制作完成。中间得到了许多人的支持  ************
'************ 在此表示感谢。感谢名单详见m208的自留地。         ************
'************                                                  ************
'************ 本类模块说明：                                   ************
'************ 名称：Tapi                                       ************
'************ 作用：本 SDK 的主要部分。详细说明参见文档。      ************
'************                                                  ************
Public oauth_consumer_key As String
Public access_token As String
Public expires_in As String
Public openid As String
Public openkey As String
Public refresh_token As String
Public state As String
Public name As String
Public nick As String
Public clientip As String
Const APIURL = "https://open.t.qq.com/api/"
Function GetAuthPage(Callback As String) As String
state = GetState()
GetAuthPage = "https://open.t.qq.com/cgi-bin/oauth2/authorize?client_id=" & oauth_consumer_key & "&response_type=token&redirect_uri=" & Callback & "&state=" & state
End Function
Function GetAccessToken(AfterCallback As String) As Boolean
access_token = GetURLParam(AfterCallback, "access_token")
expires_in = GetURLParam(AfterCallback, "expires_in")
openid = GetURLParam(AfterCallback, "openid")
openkey = GetURLParam(AfterCallback, "openkey")
refresh_token = GetURLParam(AfterCallback, "refresh_token")
If GetURLParam(AfterCallback, "state") = state Then
Debug.Print "校验成功！"
Else
Debug.Print "出错，小心csrf攻击"
Debug.Print GetURLParam(AfterCallback, "state")
Debug.Print state
GetAccessToken = False
Exit Function
End If
name = GetURLParam(AfterCallback, "name")
nick = GetURLParam(AfterCallback, "nick")
Debug.Print access_token
Debug.Print expires_in
Debug.Print openid
Debug.Print openkey
Debug.Print refresh_token
Debug.Print name
Debug.Print nick
GetAccessToken = True
End Function
Sub init(ByVal AppKey As String)
oauth_consumer_key = AppKey
Debug.Print oauth_consumer_key
clientip = GetClientIP()
End Sub
Private Function GetState() As String
For i = 0 To 9
s = s & Chr(Fix((Rnd() * 26)) + 97)
Next
GetState = s
End Function
Function CallAPI(API As String, Method As String, Param As String) As String
Dim callUrl As String
Dim Final As String
callUrl = APIURL & API
If Method = "GET" Then
callUrl = callUrl & "?oauth_consumer_key=" & oauth_consumer_key & "&access_token=" & access_token & "&openid=" & openid & "&clientip=" & clientip & "&oauth_version=2.a&scope=all&" & Param
Else
Param = "oauth_consumer_key=" & oauth_consumer_key & "&access_token=" & access_token & "&openid=" & openid & "&clientip=" & clientip & "&oauth_version=2.a&scope=all&" & Param
End If
If Method = "GET" Then
Final = XMLSend(callUrl)
Else
Final = XMLSend(callUrl, "POST", Param)
End If
If Final = "成功！" Then
CallAPI = BytesToString(XMLSendByte(), "UTF-8")
Else
CallAPI = Final
End If
End Function
Private Function GetClientIP() As String
Dim aa As Object
Dim bb As Object
Dim step1 As String
Dim step2 As String
Dim r As Variant
If XMLSend("http://iframe.ip138.com/ic.asp") = "成功！" Then
step1 = BytesToString(XMLSendByte(), "GB2312")
Set aa = CreateObject("VBscript.RegExp")
aa.Global = True
aa.IgnoreCase = True
aa.MultiLine = True
aa.Pattern = "(\d+)\.(\d+)\.(\d+)\.(\d+)"
For Each r In aa.Execute(step1)
     step2 = r
     Exit For
Next
End If
Debug.Print step2
GetClientIP = step2
End Function
Sub RefreshAccessToken()
Dim callUrl As String
Dim Final As String
state = GetState()
callUrl = "https://open.t.qq.com/cgi-bin/oauth2/access_token?client_id=" & oauth_consumer_key & "&grant_type=refresh_token&refresh_token=" & refresh_token & "&state=" & state
Final = XMLSend(callUrl)
Final = BytesToString(XMLSendByte(), "UTF-8")
access_token = GetURLParam(Final, "access_token")
expires_in = GetURLParam(Final, "expires_in")
refresh_token = GetURLParam(Final, "refresh_token")
name = GetURLParam(Final, "name")
End Sub
Private Function GetURLParam(URL As String, ParamName As String)
Dim RegExp As Object
Dim Final As Object
Set RegExp = CreateObject("VBScript.RegExp")
RegExp.Pattern = "(^|&|\\?)" & ParamName & "=([^&]*)(&|$)"
Set Final = RegExp.Execute(URL)
GetURLParam = Final(0)
GetURLParam = Replace(GetURLParam, ParamName & "=", "")
GetURLParam = Replace(GetURLParam, "&", "")
End Function
Function CallAPIPic(API As String, ParamName As Variant, ParamValue As Variant, PicPath As String, PicName As String, PicFileName As String, PicType As String)
Dim i As Integer
Dim PicUpload
Set PicUpload = New TXMLHTTPPic
Dim a As Variant
PicUpload.Charset = "UTF-8"
PicUpload.AddForm "oauth_consumer_key", oauth_consumer_key
PicUpload.AddForm "access_token", access_token
PicUpload.AddForm "openid", openid
PicUpload.AddForm "clientip", clientip
PicUpload.AddForm "oauth_version", "2.a"
PicUpload.AddForm "scope", "all"
PicUpload.AddFile PicName, PicFileName, PicType, PicPath
For i = 0 To UBound(ParamName) Step 1
a = PicUpload.AddForm(ParamName(i), ParamValue(i))
Next
CallAPIPic = PicUpload.upload(APIURL & API)
Set PicUpload = Nothing
End Function
