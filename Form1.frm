VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mostafa Server"
   ClientHeight    =   3270
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4710
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3270
   ScaleWidth      =   4710
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Clear"
      Height          =   375
      Left            =   3960
      TabIndex        =   1
      Top             =   0
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Height          =   3255
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   0
      Width           =   3975
   End
   Begin VB.Timer Timer1 
      Interval        =   80
      Left            =   2760
      Top             =   720
   End
   Begin MSWinsockLib.Winsock Winsock2 
      Index           =   0
      Left            =   4080
      Top             =   1440
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   3600
      Top             =   1920
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   80
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This code is written by Mostafa Mohamed
'email mostafahafez81@hotmail.com
'webpage www.geocities.com/ge3d
'it explaon how to make an http server which can be used in real
'to test this sample:
'1-run this server
'2-run microsoft internet explorer
'3-type http://127.0.0.1/page.htm
'then the server wil show the page
'modify it to fit your use
'so u can convert your home computer into an internet server
'if u want to test it with another friend don'tuse 127.0.0.1 use the ip in the caption of the form it is your computer ip
'so the address wil be http://(yourip)/page.htm
Private Sub Command1_Click()
Text1.Text = ""
End Sub


Private Sub Form_Load()
Dim i As Integer
Me.Caption = Winsock1.LocalIP 'Set the form caption to your server (the computer that run the program ip
Winsock1.LocalPort = 80 'set the port for the winsock to listen in
For i = 1 To 100
Load Winsock2(i) 'load the connections winsock
Winsock2(i).LocalPort = 80 + i 'settinh thier ports
Next i
Winsock1.Listen 'set the main sock to listen for incoming connectinos
End Sub

Private Sub Form_Unload(Cancel As Integer)
Winsock1.Close
End Sub

Private Sub Winsock1_ConnectionRequest(ByVal requestID As Long)
Dim i As Integer
'if some one wants to make a connection with the server
'Searching for free sock for new connection request
For i = 0 To 100
If Winsock2(i).State = 0 Then 'if the sock free(closed)
Winsock2(i).Accept requestID 'let it accept the connection ,we don't accept it by winsock1 because it is already busy in listening
Exit Sub 'exit the sub
End If
Next i
End Sub



Private Sub Winsock2_Close(Index As Integer)
'if the user closed from the browser then close the connection here
Winsock2(Index).Close
End Sub



Private Sub Winsock2_DataArrival(Index As Integer, ByVal bytesTotal As Long)
On Error GoTo errr
Dim str As String
Dim ss As Byte
Dim sc As String
Dim filen As Integer
Dim ext As String
Dim sb As String
Dim page As String
Dim i As Integer
Dim splited() As String
Dim splited2() As String
Dim splited3() As String

'when data arrived from the user browser
Winsock2(Index).GetData str, vbString 'get the data from winsock2

Text1.Text = Text1.Text + vbCrLf + str 'add the text arrived from the browser to study it(i add this line to let u see the data arrived from internet exploror or netscape to study it)

splited() = Split(str, vbCrLf) 'split it into lines "vbcrlf"= new line (as u press enter while u are typing)

For i = 0 To UBound(splited()) 'from 0 to the no of lines
splited2() = Split(splited(i), " ") 'split the line according to the space character

If i = 0 Then 'if this is the first line then
If splited2(0) <> "GET" Then Exit Sub 'if the command from the browser not get file
page = splited2(1) 'page = the second word in this line
page = Mid(page, 2, Len(page)) 'remove the slash charcter from the page name since / is used in http protocol instead of \ which we use it in windows or dos for files
splited3() = Split(page, ".") 'determine the file extenstion
ext = splited3(1)
End If

Next i 'next line



'Send the server response
str = "HTTP/1.1 200 OK"
str = str & vbCrLf & "Server: Mostafa-Server" '"Server: Microsoft-IIS/5.0"
str = str & vbCrLf & "Date: " & Format(Date, "Medium Date", vbMonday, vbFirstJan1)
str = str & vbCrLf & "Content-Type: text/html"
str = str & vbCrLf & "Accept-Ranges: bytes"
str = str & vbCrLf & "Last-Modifid " & FileDateTime(App.Path & "\" & page)

'If the file needed by the browser is text then send it
If ext = "txt" Or ext = "TXT" Or ext = "htm" Or ext = "html" Or ext = "HTML" Or ext = "htm" Then 'determine the extenstion
filen = FreeFile
Open App.Path & "\" & page For Input As #filen 'open it for input
Do
Input #1, sc
sb = sb + sc 'fill the data
Loop Until EOF(1)
Close #filen
End If

'If the file needed by the browser is binary then send it
If ext = "jpg" Or ext = "JPG" Or ext = "zip" Or ext = "ZIP" Or ext = "gif" Or ext = "GIF" Then
filen = FreeFile
Open App.Path & "\" & page For Binary As #filen
Do
Get #filen, , ss
sb = sb & Chr(ss) 'fill the data in sb to be sent
Loop Until EOF(1)
Close #filen
End If

str = str & vbCrLf & "Content-Length: " & Len(sb)
str = str & vbCrLf & ""
str = str & vbCrLf & sb
Debug.Print str 'this line to let u see the data sent by this server to ie or netscape in the debug window
Winsock2(Index).SendData str 'send it
Exit Sub
errr: 'if error
'then send an error page
filen = FreeFile
Open App.Path & "\error.htm" For Input As #filen
Do
Input #1, sc
sb = sb + sc
Loop Until EOF(1)
Close #filen
str = str & vbCrLf & "Content-Length: " & Len(sb) 'calculate the page size
str = str & vbCrLf & ""
str = str & vbCrLf & sb
Winsock2(Index).SendData str
Debug.Print Err.Description
End Sub

Private Sub Winsock2_SendComplete(Index As Integer)
Winsock2(Index).Close
End Sub
