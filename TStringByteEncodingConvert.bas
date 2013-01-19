Attribute VB_Name = "TStringByteEncodingConvert"
Option Explicit

'************ Tencent Weibo SDK for Visual Basic 6 ( OAuth 2 ) ************
'************                                                  ************
'************ 此 SDK 由 m208 制作完成。中间得到了许多人的支持  ************
'************ 在此表示感谢。感谢名单详见m208的自留地。         ************
'************                                                  ************
'************ 本模块说明：                                     ************
'************ 名称：TStringByteEncodingConvert                 ************
'************ 作用：负责字符与字节以及编码之间的转换。此模块来 ************
'************ 自CSDN：http://bbs.csdn.net/topics/250072337     ************
'************                                   在此表示感谢。 ************

Public Const adTypeBinary = 1
Public Const adTypeText = 2
Public Const adLongVarBinary = 205

'字节数组转指定字符集的字符串
Public Function BytesToString(vtData, ByVal strCharset)
    Dim objFile
    Set objFile = New ADODB.Stream
    objFile.Type = adTypeBinary
    objFile.Open
    If VarType(vtData) = vbString Then
        objFile.Write BinaryToBytes(vtData)
    Else
        objFile.Write vtData
    End If
    objFile.Position = 0
    objFile.Type = adTypeText
    objFile.Charset = strCharset
    BytesToString = objFile.ReadText(-1)
    objFile.Close
    Set objFile = Nothing
End Function

'字节字符串转字节数组，即经过MidB/LeftB/RightB/ChrB等处理过的字符串
Public Function BinaryToBytes(vtData)
    Dim rs
    Dim lSize
    lSize = LenB(vtData)
    Set rs = New ADODB.Recordset
    rs.fields.Append "Content", adLongVarBinary, lSize
    rs.Open
    rs.AddNew
    rs("Content").AppendChunk vtData
    rs.Update
    BinaryToBytes = rs("Content").GetChunk(lSize)
    rs.Close
    Set rs = Nothing
End Function

'指定字符集的字符串转字节数组
Public Function StringToBytes(ByVal strData, ByVal strCharset)
    Dim objFile
    Set objFile = New ADODB.Stream
    objFile.Type = adTypeText
    objFile.Charset = strCharset
    objFile.Open
    objFile.WriteText strData
    objFile.Position = 0
    objFile.Type = adTypeBinary
    If UCase(strCharset) = "UNICODE" Then
        objFile.Position = 2 'delete UNICODE BOM
    ElseIf UCase(strCharset) = "UTF-8" Then
        objFile.Position = 3 'delete UTF-8 BOM
    End If
    StringToBytes = objFile.Read(-1)
    objFile.Close
    Set objFile = Nothing
End Function

'获取文件内容的字节数组
Public Function GetFileBinary(ByVal strPath)
    Dim objFile
    Set objFile = New ADODB.Stream
    objFile.Type = adTypeBinary
    objFile.Open
    objFile.LoadFromFile strPath
    GetFileBinary = objFile.Read(-1)
    objFile.Close
    Set objFile = Nothing
End Function


Function URLEncodeUTF8(szInput As String) As String
Dim wch, uch, szRet
Dim x
Dim nAsc, nAsc2, nAsc3

If szInput = "" Then
    Exit Function
End If

For x = 1 To Len(szInput)
    wch = Mid(szInput, x, 1)
    nAsc = AscW(wch)

    If nAsc < 0 Then nAsc = nAsc + 65536

    If (nAsc And &HFF80) = 0 Then
        szRet = szRet & wch
    Else
        If (nAsc And &HF000) = 0 Then
            uch = "%" & Hex(((nAsc \ 2 ^ 6)) Or &HC0) & "%" & Hex(nAsc And &H3F Or &H80)
            szRet = szRet & uch
        Else
            uch = "%" & Hex((nAsc \ 2 ^ 12) Or &HE0) & "%" & _
            Hex((nAsc \ 2 ^ 6) And &H3F Or &H80) & "%" & _
            Hex(nAsc And &H3F Or &H80)
            szRet = szRet & uch
        End If
    End If
Next

szRet = Replace$(szRet, " ", "%20")

URLEncodeUTF8 = szRet
End Function

