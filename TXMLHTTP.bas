Attribute VB_Name = "TXMLHTTP"
Option Explicit

'************ Tencent Weibo SDK for Visual Basic 6 ( OAuth 2 ) ************
'************                                                  ************
'************ 此 SDK 由 m208 制作完成。中间得到了许多人的支持  ************
'************ 在此表示感谢。感谢名单详见m208的自留地。         ************
'************                                                  ************
'************ 本模块说明：                                     ************
'************ 名称：TXMLHTTP                                   ************
'************ 作用：负责普通的GET和POST请求。此模块来自CSDN：  ************
'************ http://download.csdn.net/download/fisheep_works/ ************
'************ 1424456                           在此表示感谢。 ************

Public XMLSendByte() As Byte
'需要引用 xml
'xml发送数据
Function XMLSend(ByVal URL As String, Optional Method = "Get", Optional ByVal Form As String = "") As String
Dim Http As New xmlHttp
'On Error GoTo RErr

'使用正确的编码进行转换
URL = URLEncodeUTF8(URL)
If UCase(Method) = "POST" And Form <> "" Then
    Form = URLEncodeUTF8(Form)
Else
    Form = ""
End If

Http.Open Method, URL, False
Http.setRequestHeader "Accept", "*/*"
Http.setRequestHeader "Accept-Language", "zh-cn"
Http.setRequestHeader "User-Agent", "Mozilla/5.0 (Windows NT 6.2; rv:21.0) Gecko/20130116 Firefox/21.0" '您也可以自定义下
Http.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"


'如无法访问网络则会出现错误
If Form = "" Then
    Http.send
Else
    Http.send Form
End If

If Http.ReadyState = 4 And Http.Status = 200 Then
    XMLSendByte() = Http.responseBody
    XMLSend = "成功！"
Else
    XMLSend = "出现错误:readyState:" & Http.ReadyState & "Status:" & Http.Status
End If

Exit Function
RErr:
XMLSend = "出现错误:" & Err.Description
End Function
