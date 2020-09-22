Attribute VB_Name = "mod_HTTP"
Option Explicit

Public Function ParseHeader(ByVal strExpression As String, _
                            ByVal strVariable As String, _
                            Optional optDelimiter As String = vbCrLf) _
                            As String
    Dim I As Integer
    Dim strTemp As String
    
    If Not Right(Trim(strVariable), 1) = ":" Then
        strVariable = strVariable & ": "
    End If
    
    I = InStr(1, strExpression, strVariable, vbTextCompare)
    If I = 0 Then
        ParseHeader = vbNullString
    Else
        strTemp = Mid(strExpression, I + Len(strVariable))
        I = InStr(1, strTemp, vbCrLf, vbTextCompare)
        If I = 0 Then
            ParseHeader = vbNullString
        Else
            ParseHeader = Left(strTemp, I - 1)
        End If
    End If
End Function

Public Function ParseUrl(ByVal strUrl As String, _
                          ByRef strHost As String, _
                          ByRef lngPort As Long, _
                          ByRef strRequest As String)
    Dim I As Integer
    Dim X As Integer
    Dim strSub As String
    
    If LCase(Left(strUrl, Len("http://"))) = "http://" Then
        'http:// exists?
        strUrl = Mid(strUrl, Len("http://") + 1)
    End If
    
    'www. is at some sites important for unknown reason
    '--------------------------------------------------
    'If LCase(Left(strUrl, Len("www."))) = "www." Then
    '    'www. exists?
    '    strUrl = Mid(strUrl, Len("www.") + 1)
    'End If
    
    I = InStr(strUrl, "/")
    If Not I = 0 Then
        strRequest = Mid(strUrl, I)
        strSub = Left(strUrl, I - 1)
        I = InStr(strSub, ":")
        If Not I = 0 Then
            lngPort = Val(Mid(strSub, I + 1))
            strHost = Left(strSub, I - 1)
        Else
            lngPort = 80
            strHost = strSub
        End If
    Else
        strRequest = "/"
        I = InStr(strUrl, ":")
        If Not I = 0 Then
            lngPort = Val(Mid(strUrl, I + 1))
            strHost = Left(strUrl, I - 1)
        Else
            lngPort = 80
            strHost = strUrl
        End If
    End If
End Function
