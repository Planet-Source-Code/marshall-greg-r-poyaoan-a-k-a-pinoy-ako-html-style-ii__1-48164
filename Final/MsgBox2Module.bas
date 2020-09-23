Attribute VB_Name = "MsgBox2Module"
'........THIS IS THE MODULE OF MsgBox2 by: Pinoy Ako!
'This is just a sample and sort of a test
'wether you like the MsgBox2 code or not, so PLEASE COMMENT.
'I did not improve it much right now
'because my focus is coding the HTML Style II.
'Maybe later this month I will release new version
'of MsgBox2.


'........HOW to make MsgBox2 work?
'to make the MsgBox2 work, picMsg(picture box) and a timer must exist.
'inside the timer paste this -> MsgBoxWpicture
'And the form must be named Form1.
'But you can always change the property name of
'picMsg,Form1. But by changing them you have to
'edit code in Public Sub MsgBoxWpicture().

'........Using MsgBox2
'...Example
'Private Sub command1_Click()
'Dim x as integer
'Dim y as integer
'Dim width as integer
'Dim height as integer
'............'  Picture x and y coordinates.Picture width and height  '
'x=0
'y=0
'width=100
'height=100
'Call MsgBox2(Form1.hwnd,"Message: Hello","Title",vbOKOnly,x,y,width,height)
'End Sub
'

'.........Some TIPS
'you can also make the message box Elliptic, Round Rect ,etc.. by using API's SetwindowRgn.
'you can also put icon in the message box by using API's DrawIcon.
'you can also put text in the messgae box by using API's TextOut,SetbkMode.
'you can also change the buttons text by using API's SetWindowText.
'maybe more. But for now "that all folks".

'API's Declare
Public Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Public Declare Function MessageBox Lib "user32" Alias "MessageBoxA" (ByVal hwnd As Long, ByVal lpText As String, ByVal lpCaption As String, ByVal wType As Long) As Long
Public Declare Function GetForegroundWindow Lib "user32" () As Long
Public Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal ByteLen As Long)
Public Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Public Declare Function SetWindowText Lib "user32" Alias "SetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String) As Long



Dim Widthm As Integer
Dim Heightm As Integer
Dim Xm As Integer
Dim Ym As Integer
Dim Titlem As String
Dim bwnd As Long
Dim Ret As Long

'msgbox2 function
Public Function MsgBox2(ByVal hwndPic As Long, _
    Optional ByVal Prompt As String, Optional ByVal Title As String = "  ", _
    Optional ByVal Buttons As VbMsgBoxStyle = vbOKOnly, _
    Optional ByVal PicX As Integer = 0, _
    Optional ByVal PicY As Integer = 0, _
    Optional ByVal PicWidth As Integer = 60, _
    Optional ByVal PicHeight As Integer = 60)
        On Error Resume Next
        Xm = PicX
        Ym = PicY
        Titlem = Title
        Widthm = PicWidth
        Heightm = PicHeight
        MessageBox hwndPic, Prompt, Title, Buttons
        
End Function

'put picture in message box
Public Sub MsgBoxWPicture()
On Error Resume Next

    'find message box
    Ret = FindWindow("#32770", Titlem)
        If CBool(Ret) = True Then
                'Get Ret device context
                Ret = GetDC(Ret)
                'Set the graphics mode to persistent
                Form1.picMsg.AutoRedraw = True
                'API uses pixels
                Form1.picMsg.ScaleMode = vbPixels
                'copy picture from picMsg(picture box)
                StretchBlt Ret, Xm, Ym, Widthm, Heightm, Form1.picMsg.hdc, 0, 0, Form1.picMsg.ScaleWidth, Form1.picMsg.ScaleHeight, vbSrcCopy
                DoEvents
           
        End If

End Sub




