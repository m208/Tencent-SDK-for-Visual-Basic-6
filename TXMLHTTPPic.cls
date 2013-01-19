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
Option Explicit 'ǿ�Ʊ�������

'************ Tencent Weibo SDK for Visual Basic 6 ( OAuth 2 ) ************
'************                                                  ************
'************ �� SDK �� m208 ������ɡ��м�õ��������˵�֧��  ************
'************ �ڴ˱�ʾ��л����л�������m208�������ء�         ************
'************                                                  ************
'************ ����ģ��˵����                                   ************
'************ ���ƣ�TXMLHTTPPic                                ************
'************ ���ã��������ͼƬ��POST���󡣴���ģ������CSDN�� ************
'************ http://bbs.csdn.net/topics/250072337             ************
'************                                   �ڴ˱�ʾ��л�� ************

Public xmlHttp 'XMLHTTP����
Public objTemp 'ADODB.Stream �����ı����ֽ�
Public strCharset, strBoundary '�ַ������Boundary��multipart/form-data��ʽ�зָ����ݵķָ�����

Private Sub Class_Initialize() 'ģ���ʼ��
    Set xmlHttp = New MSXML2.XMLHTTP30 '����ΪXMLHTTP����
    Set objTemp = New ADODB.Stream '����ΪADODB.Stream����
    objTemp.Type = adTypeBinary '�Դ��ֽ�ģʽ
    objTemp.Open '��
    strCharset = "UTF-8" '�����ַ�����
    strBoundary = GetBoundary() '����Boundary
End Sub

Private Sub Class_Terminate() '�����¼�����������
    objTemp.Close
    Set objTemp = Nothing
    Set xmlHttp = Nothing
End Sub

'��ȡ�Զ���ı������ݷֽ���
Private Function GetBoundary()
    Dim ret(24)
    Dim table
    Dim i
    table = "ABCDEFGHIJKLMNOPQRSTUVWXZYabcdefghijklmnopqrstuvwxzy0123456789" '�ַ���
    Randomize '���������
    For i = 0 To UBound(ret)
        ret(i) = Mid(table, Int(Rnd() * Len(table) + 1), 1)
    Next
    GetBoundary = "__NextPart__ " & Join(ret, Empty)
End Function

'�����ϴ�ʹ�õ��ַ���
Public Property Let Charset(ByVal strValue)
    strCharset = strValue
End Property

'�����ı�������ƺ�ֵ
Public Sub AddForm(ByVal strName, ByVal strValue)
    Dim tmp
    tmp = "\r\n--$1\r\nContent-Disposition: form-data; name=""$2""\r\n\r\n$3"
    tmp = Replace(tmp, "\r\n", vbCrLf) '�滻���з�
    tmp = Replace(tmp, "$1", strBoundary) '�滻Boundary
    tmp = Replace(tmp, "$2", strName) '�滻�ı�������
    tmp = Replace(tmp, "$3", strValue) '�滻�ı���ֵ
    objTemp.Write StringToBytes(tmp, strCharset)
End Sub

'�����ļ��������/�ļ�����/�ļ�MIME����/�ļ�·�����ļ��ֽ�����
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

'����multipart/form-data�������
Private Sub AddEnd()
    Dim tmp
    tmp = Replace("\r\n--$1--\r\n", "$1", strBoundary)
    tmp = Replace(tmp, "\r\n", vbCrLf) 'PS:ԭ�������Ǽ�����ˣ��������һ��ΪͼƬʱ����ִ���
    objTemp.Write StringToBytes(tmp, strCharset)
    objTemp.Position = 0
End Sub

'�ϴ���ָ����URL�������ط�����Ӧ��
Public Function upload(ByVal strURL)
    Call AddEnd
    xmlHttp.Open "POST", strURL, False
xmlHttp.setRequestHeader "Content-Type", "multipart/form-data; boundary=" & strBoundary
    xmlHttp.setRequestHeader "Content-Length", objTemp.Size
    xmlHttp.send objTemp
    upload = BytesToString(xmlHttp.responseBody, strCharset)
End Function
