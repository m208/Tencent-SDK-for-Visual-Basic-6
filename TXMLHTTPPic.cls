VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TXMLHTTPPic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit '强制变量声明

'************ Tencent Weibo SDK for Visual Basic 6 ( OAuth 2 ) ************
'************                                                  ************
'************ 此 SDK 由 m208 制作完成。中间得到了许多人的支持  ************
'************ 在此表示感谢。感谢名单详见m208的自留地。         ************
'************                                                  ************
'************ 本类模块说明：                                   ************
'************ 名称：TXMLHTTPPic                                ************
'************ 作用：负责包含图片的POST请求。此类模块来自CSDN： ************
'************ http://bbs.csdn.net/topics/250072337             ************
'************                                   在此表示感谢。 ************

Public xmlHttp 'XMLHTTP对象
Public objTemp 'ADODB.Stream 处理文本和字节
Public strCharset, strBoundary '字符编码和Boundary（multipart/form-data格式中分割数据的分隔符）

Private Sub Class_Initialize() '模块初始化
    Set xmlHttp = New MSXML2.XMLHTTP30 '设置为XMLHTTP对象
    Set objTemp = New ADODB.Stream '设置为ADODB.Stream对象
    objTemp.Type = adTypeBinary '对待字节模式
    objTemp.Open '打开
    strCharset = "UTF-8" '设置字符编码
    strBoundary = GetBoundary() '生成Boundary
End Sub

Private Sub Class_Terminate() '销毁事件，不作解释
    objTemp.Close
    Set objTemp = Nothing
    Set xmlHttp = Nothing
End Sub

'获取自定义的表单数据分界线
Private Function GetBoundary()
    Dim ret(24)
    Dim table
    Dim i
    table = "ABCDEFGHIJKLMNOPQRSTUVWXZYabcdefghijklmnopqrstuvwxzy0123456789" '字符池
    Randomize '随机数种子
    For i = 0 To UBound(ret)
        ret(i) = Mid(table, Int(Rnd() * Len(table) + 1), 1)
    Next
    GetBoundary = "__NextPart__ " & Join(ret, Empty)
End Function

'设置上传使用的字符集
Public Property Let Charset(ByVal strValue)
    strCharset = strValue
End Property

'添加文本域的名称和值
Public Sub AddForm(ByVal strName, ByVal strValue)
    Dim tmp
    tmp = "\r\n--$1\r\nContent-Disposition: form-data; name=""$2""\r\n\r\n$3"
    tmp = Replace(tmp, "\r\n", vbCrLf) '替换换行符
    tmp = Replace(tmp, "$1", strBoundary) '替换Boundary
    tmp = Replace(tmp, "$2", strName) '替换文本域名称
    tmp = Replace(tmp, "$3", strValue) '替换文本域值
    objTemp.Write StringToBytes(tmp, strCharset)
End Sub

'设置文件域的名称/文件名称/文件MIME类型/文件路径或文件字节数组
Public Sub AddFile(ByVal strName, ByVal strFileName, ByVal strFileType, vtValue)
    Dim tmp
    tmp = "\r\n--$1\r\nContent-Disposition: form-data; name=""$2""; filename=""$3""\r\nContent-Type: $4\r\n\r\n"
    tmp = Replace(tmp, "\r\n", vbCrLf)
    tmp = Replace(tmp, "$1", strBoundary)
    tmp = Replace(tmp, "$2", strName)
    tmp = Replace(tmp, "$3", strFileName)
    tmp = Replace(tmp, "$4", strFileType)
    objTemp.Write StringToBytes(tmp, strCharset)
    If VarType(vtValue) = (vbByte Or vbArray) Then
        objTemp.Write vtValue
    Else
        objTemp.Write GetFileBinary(vtValue)
    End If
End Sub

'设置multipart/form-data结束标记
Private Sub AddEnd()
    Dim tmp
    tmp = Replace("\r\n--$1--\r\n", "$1", strBoundary)
    tmp = Replace(tmp, "\r\n", vbCrLf) 'PS:原作者忘记加这个了，导致最后一项为图片时会出现错误
    objTemp.Write StringToBytes(tmp, strCharset)
    objTemp.Position = 0
End Sub

'上传到指定的URL，并返回服务器应答
Public Function upload(ByVal strURL)
    Call AddEnd
    xmlHttp.Open "POST", strURL, False
xmlHttp.setRequestHeader "Content-Type", "multipart/form-data; boundary=" & strBoundary
    xmlHttp.setRequestHeader "Content-Length", objTemp.Size
    xmlHttp.send objTemp
    upload = BytesToString(xmlHttp.responseBody, strCharset)
End Function

