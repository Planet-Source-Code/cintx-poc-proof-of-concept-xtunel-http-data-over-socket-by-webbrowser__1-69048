VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frm_Main 
   Caption         =   "xTunel"
   ClientHeight    =   6555
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8535
   LinkTopic       =   "Form1"
   ScaleHeight     =   6555
   ScaleWidth      =   8535
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmd_Browse 
      Caption         =   "Go"
      Default         =   -1  'True
      Height          =   255
      Left            =   7440
      TabIndex        =   1
      Top             =   15
      Width           =   1095
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   6300
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin MSWinsockLib.Winsock Socket 
      Left            =   600
      Top             =   1560
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.TextBox txt_Url 
      Height          =   285
      Left            =   0
      TabIndex        =   0
      Text            =   "http://yahoo.com"
      Top             =   0
      Width           =   7455
   End
   Begin SHDocVwCtl.WebBrowser wBrowser 
      Height          =   5295
      Left            =   0
      TabIndex        =   3
      Top             =   270
      Width           =   8535
      ExtentX         =   15055
      ExtentY         =   9340
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
End
Attribute VB_Name = "frm_Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type tBrowse
    strRequest As String
    strHost As String
    lngPort As Long
    bolPost As Boolean
    strPost As String
    bolHead As Boolean
    lngRec As Long
    lngSize As Long
    strData As String
End Type
Dim bolBrowse As Boolean
Dim wHTML As IHTMLDocument2

Dim pBrowse As tBrowse

Private Sub cmd_Browse_Click()
    ParseUrl txt_Url, pBrowse.strHost, pBrowse.lngPort, pBrowse.strRequest
    If Not pBrowse.lngPort = 0 Then
        StatusBar1.Panels(1).Text = "Status: Connecting to " & pBrowse.strHost
        Socket.Close
        Socket.Connect pBrowse.strHost, pBrowse.lngPort
    End If
End Sub

Private Sub Form_Load()
    wBrowser.Navigate "about:blank"
    wBrowser.Silent = True
    MsgBox "Make sure you compile this EXE before you start it" & vbCrLf & _
            "The WebBrowser control dont like the VB IDE ;P", vbOKOnly, "Note!"
End Sub

Private Sub Form_Resize()
    On Error GoTo exError
    StatusBar1.Panels(1).Width = Me.ScaleWidth - 1000
    StatusBar1.Panels(2).Width = 1000
    wBrowser.Width = Me.ScaleWidth
    wBrowser.Height = Me.ScaleHeight - StatusBar1.Height - 300
    txt_Url.Width = Me.ScaleWidth - cmd_Browse.Width
    cmd_Browse.Left = txt_Url.Width
exError:
End Sub

Private Sub Socket_Close()
    'Connection Closed
    StatusBar1.Panels(1).Text = "Status: Complete"
    StatusBar1.Panels(2).Text = "100%"
    
    'InsertUrls txt_Url, pBrowse.strData
    'InsertUrls - Doesnt work proper
    
    wBrowser.Document.write pBrowse.strData
    
    pBrowse.strData = ""
End Sub

Private Sub Socket_Connect()
    Dim strHead As String
    
    If pBrowse.bolPost = True Then
        'Post Packet
        
    Else
        'Get Packet
        strHead = "GET " & pBrowse.strRequest & " HTTP/1.1" & vbCrLf & _
                    "Host: " & pBrowse.strHost & ":" & pBrowse.lngPort & vbCrLf & _
                    "Accept: */*" & vbCrLf & _
                    "Connection: close" & vbCrLf & _
                    "User-Agent: xTunel/1.0" & vbCrLf & _
                     vbCrLf
    End If
    
    'Reset Variables
    pBrowse.lngRec = 0
    pBrowse.lngSize = 0
    pBrowse.bolHead = False
        
    wBrowser.Navigate "about:blank"
        
    If Socket.State = sckConnected Then
        'Make sure connection is alive
        Socket.SendData strHead
    End If
End Sub

Private Sub Socket_DataArrival(ByVal bytesTotal As Long)
    Dim strData As String
    
    If Not bytesTotal = 0 Then
        'We received not 0 bytes
        
        If Socket.State = sckConnected Then
            Socket.GetData strData, vbString, bytesTotal
            
            ProcessData strData
        End If
    End If
End Sub

Sub ProcessData(ByVal strData As String)
    If LCase(Left(strData, 4)) = "http" Then
        ProcessHeader strData
    Else 'Usual Data
        'Update Length
        pBrowse.lngRec = pBrowse.lngRec + Len(strData)
        pBrowse.strData = pBrowse.strData & strData
    End If
End Sub

Sub ProcessHeader(ByVal strData As String)
    Dim strHeader As String
    Dim strStatus() As String
    Dim I As Integer
    
    If pBrowse.bolHead = False Then
        'Header was not received yet
        pBrowse.bolHead = True
        I = InStr(strData, vbCrLf & vbCrLf)
        If Not I = 0 Then
            strHeader = Left(strData, I - 2)
            strStatus = Split(strHeader, " ")
            strData = Mid(strData, I + 4)
            If (strStatus(1) = "200") Then
                '200 OK
                '------
                pBrowse.lngSize = Val(ParseHeader(strHeader, "Content-Length"))
                Status_200 strData
            ElseIf (strStatus(1) = "302") Then
                '302 Movied
                '----------
                Status_304 strHeader
            ElseIf (strStatus(1) = "301") Then
                '302 Movied
                '----------
                Status_304 strHeader
            ElseIf (strStatus(1) = "404") Then
                '404 Not found
                '----------
                
            End If
        End If
    Else 'Usual Data
        'Update Length
        DataReceive strData
    End If
    If pBrowse.bolHead = True Then
        UpdateStatus_Rec
    End If
End Sub

Sub DataReceive(ByVal strData As String)
    StatusBar1.Panels(1).Text = "Status: Loading"
    pBrowse.lngRec = pBrowse.lngRec + Len(strData)
    pBrowse.strData = pBrowse.strData & strData
End Sub

Sub UpdateStatus_Rec()
    Dim lngComplete As Integer
    If Not pBrowse.lngSize = 0 Then
        lngComplete = pBrowse.lngRec / pBrowse.lngSize
        lngComplete = lngComplete * 100
        StatusBar1.Panels(2).Text = Format(lngComplete, "0") & "%"
    Else
        StatusBar1.Panels(2).Text = "?"
    End If
End Sub

Sub Status_304(ByVal strHeader As String)
    Dim strLocation As String
    strLocation = ParseHeader(strHeader, "Location")
    If Not strLocation = vbNullString Then
        StatusBar1.Panels(1).Text = "Status: Redirection"
        ParseUrl strLocation, pBrowse.strHost, pBrowse.lngPort, pBrowse.strRequest
        If Not pBrowse.lngPort = 0 Then
            txt_Url = strLocation
            StatusBar1.Panels(1).Text = "Status: Connecting to " & pBrowse.strHost
            Socket.Close
            Socket.Connect pBrowse.strHost, pBrowse.lngPort
        End If
    End If
End Sub

Sub Status_200(ByVal strData As String)
    StatusBar1.Panels(1).Text = "Status: Page found"
    
    'Existing Data
    pBrowse.lngRec = Len(strData)

    'Display existing data
    pBrowse.strData = pBrowse.strData & strData
End Sub

Private Sub Socket_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    StatusBar1.Panels(1).Text = "Status: Unable to connect " & pBrowse.strHost
End Sub

Private Sub wBrowser_BeforeNavigate2(ByVal pDisp As Object, URL As Variant, Flags As Variant, TargetFrameName As Variant, PostData As Variant, Headers As Variant, Cancel As Boolean)
    If Not URL = "about:blank" Then
        Cancel = 1
        ParseUrl URL, pBrowse.strHost, pBrowse.lngPort, pBrowse.strRequest
        If Not pBrowse.lngPort = 0 Then
            txt_Url = URL
            StatusBar1.Panels(1).Text = "Status: Connecting to " & pBrowse.strHost
            Socket.Close
            Socket.Connect pBrowse.strHost, pBrowse.lngPort
        End If
    End If
End Sub

