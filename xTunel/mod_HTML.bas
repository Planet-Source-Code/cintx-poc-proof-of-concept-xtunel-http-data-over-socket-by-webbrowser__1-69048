Attribute VB_Name = "mod_HTML"
Option Explicit

Public Function InsertUrls(ByVal strUrl As String, ByRef strData As String)
    Dim I As Long
    Dim X As Long
    Dim strNew As String
    Dim strLeft As String
    Dim strRight As String
    Dim strTemp As String
    
    If Right(strUrl, 1) = "/" Then
        strUrl = Mid(strUrl, 1, Len(strUrl) - 1)
    End If
    
    I = InStr(1, strData, "href=", vbTextCompare)
    Debug.Print I
    Do While Not I = 0
        strLeft = Left(strData, I - 1)
        strTemp = Mid(strData, I + Len("href="))
        If Left(strTemp, 1) = Chr(34) Then
            strTemp = Mid(strTemp, 2)
        End If
        X = InStr(strTemp, Chr(34))
        If Not X = 0 Then
            strTemp = Left(strTemp, X - 1)
            strRight = Mid(strTemp, X + 1)
        Else
            X = InStr(strTemp, " ")
            If Not X = 0 Then
                strTemp = Left(strTemp, X - 1)
                strRight = Mid(strTemp, X + 1)
            End If
        End If
        If Not LCase(Left(strTemp, 4)) = "http" And Not LCase(Left(strTemp, 4)) = "www." Then
            If Left(strTemp, 1) = "/" Then
                strTemp = strUrl & strTemp
            Else
                strTemp = strUrl & "/" & strTemp
            End If
            strData = strLeft & "href=" & Chr(34) & strTemp & Chr(34) & strRight
        End If
        I = InStr(I + 1, strData, "href=", vbTextCompare)
    Loop
    
    I = InStr(1, strData, "src=", vbTextCompare)
    Debug.Print I
    Do While Not I = 0
        strLeft = Left(strData, I - 1)
        strTemp = Mid(strData, I + Len("src="))
        If Left(strTemp, 1) = Chr(34) Then
            strTemp = Mid(strTemp, 2)
        End If
        X = InStr(strTemp, Chr(34))
        If Not X = 0 Then
            strTemp = Left(strTemp, X - 1)
            strRight = Mid(strTemp, X + 1)
        Else
            X = InStr(strTemp, " ")
            If Not X = 0 Then
                strTemp = Left(strTemp, X - 1)
                strRight = Mid(strTemp, X + 1)
            End If
        End If
        If Not LCase(Left(strTemp, 4)) = "http" And Not LCase(Left(strTemp, 4)) = "www." Then
            If Left(strTemp, 1) = "/" Then
                strTemp = strUrl & strTemp
            Else
                strTemp = strUrl & "/" & strTemp
            End If
            strData = strLeft & "src=" & Chr(34) & strTemp & Chr(34) & strRight
        End If
        I = InStr(I + 1, strData, "src=", vbTextCompare)
    Loop
End Function
