Attribute VB_Name = "TXMLHTTP"
Option Explicit

'************ Tencent Weibo SDK for Visual Basic 6 ( OAuth 2 ) ************
'************                                                  ************
'************ �� SDK �� m208 ������ɡ��м�õ�������˵�֧��  ************
'************ �ڴ˱�ʾ��л����л�������m208�������ء�         ************
'************                                                  ************
'************ ��ģ��˵����                                     ************
'************ ���ƣ�TXMLHTTP                                   ************
'************ ���ã�������ͨ��GET��POST���󡣴�ģ������CSDN��  ************
'************ http://download.csdn.net/download/fisheep_works/ ************
'************ 1424456                           �ڴ˱�ʾ��л�� ************

Public XMLSendByte() As Byte
'��Ҫ���� xml
'xml��������
Function XMLSend(ByVal URL As String, Optional Method = "Get", Optional ByVal Form As String = "") As String
Dim Http As New xmlHttp
'On Error GoTo RErr

'ʹ����ȷ�ı������ת��
URL = URLEncodeUTF8(URL)
If UCase(Method) = "POST" And Form <> "" Then
    Form = URLEncodeUTF8(Form)
Else
    Form = ""
End If

Http.Open Method, URL, False
Http.setRequestHeader "Accept", "*/*"
Http.setRequestHeader "Accept-Language", "zh-cn"
Http.setRequestHeader "User-Agent", "Mozilla/5.0 (Windows NT 6.2; rv:21.0) Gecko/20130116 Firefox/21.0" '��Ҳ�����Զ�����
Http.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"


'���޷��������������ִ���
If Form = "" Then
    Http.send
Else
    Http.send Form
End If

If Http.ReadyState = 4 And Http.Status = 200 Then
    XMLSendByte() = Http.responseBody
    XMLSend = "�ɹ���"
Else
    XMLSend = "���ִ���:readyState:" & Http.ReadyState & "Status:" & Http.Status
End If

Exit Function
RErr:
XMLSend = "���ִ���:" & Err.Description
End Function
