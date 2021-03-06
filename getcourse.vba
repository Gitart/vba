200?'200px':''+(this.scrollHeight+5)+'px');">Option Explicit
Dim sHtmlCode$, sURL$
Sub tt()
Dim strTextA, strTextB
sURL = "http://zukkerro.ru/TST_GZ/data_EUR.txt"
URL2HTML
strTextA = Split(sHtmlCode$, vbNewLine)
strTextB = Split(strTextA(3), ";", 4)
MsgBox strTextB(1)
MsgBox strTextB(2)
End Sub
Private Sub URL2HTML()
'Загружает Web-страницу, заданную переменной sURL, и помещает HTML в sHtmlCode
Dim objHttp As Object
On Error Resume Next
Set objHttp = CreateObject("MSXML2.XMLHTTP.3.0")
If Err.Number <> 0 Then
Err.Clear
Set objHttp = CreateObject("MSXML2.XMLHTTP")
If Err.Number <> 0 Then
Set objHttp = CreateObject("MSXML.XMLHTTPRequest")
End If
End If
If objHttp Is Nothing Then
MsgBox "Невозможно создать объект для подключения к интернет!", 48, "Ошибка"
End
End If
If objHttp Is Nothing Then Exit Sub
objHttp.Open "GET", sURL, False
On Error Resume Next
objHttp.Send
If Err.Number <> 0 Then
MsgBox "Отсутствует доступ в интернет!", 48, "Ошибка"
End
End If
On Error GoTo 0
sHtmlCode = objHttp.responseText
Set objHttp = Nothing
End Sub
