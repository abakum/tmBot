#Const Obj = True
#If Obj Then
 Const adTypeBinary = 1
 Const adTypeText = 2
 Const adModeReadWrite = 3
 Const adSaveCreateOverWrite = 2
#End If
Const api = "https://api.telegram.org/bot"
#Const Deb = 1 'limits send per second and do debug.print
Const logFile = "body.txt" 'if Deb=2 then save body to logfile

#If Obj Then
Public Function tmBotSend(token As String, chat_id As String, Optional text As String = "", Optional param As Object) As String
 Dim d As Object
#Else
Public Function tmBotSend(token As String, chat_id As String, Optional text As String = "", Optional param As Dictionary) As String
 Dim d As Dictionary
#End If
 'https://www.planetaexcel.ru/forum/index.php?PAGE_NAME=message&FID=1&TID=93149&TITLE_SEO=93149-kak-sdelat-otpravku-v-telegram-iz-makrosa-vba-excel&MID=1193376#message1193376
 'use ToD(key, value, key2, value2, ...) to setup param
 Set d = ToD("chat_id", chat_id)
 Set medias = New Collection
 If param Is Nothing Then
  If Len(text) Then
   d.Add "text", text
   tmBotSend = tmBot(token, "sendMessage", d)
  End If
  Exit Function
 Else
  If param.Count = 0 Then
   If Len(text) Then
    d.Add "text", text
    tmBotSend = tmBotForm(token, "sendMessage", d)
   End If
   Exit Function
  End If
  filename = param.Keys
  pavd = param.Items  'pavd as "photo", "animation", "audio", "voice", "video", "video_note" and "document" as default
 End If
 sendChatAction = True 'only one sendChatAction
 For i = 0 To UBound(filename)
  send = "document"
  dfn = "" 'filename w/o path case exist
  On Error Resume Next
  If Len(filename(i)) Then dfn = Dir(filename(i))
  On Error GoTo 0
  Select Case LCase(pavd(i))
  Case "photo", "animation", "audio", "voice", "video", "video_note"
   send = LCase(pavd(i))
  Case Else
   If Len(dfn) Then
    dfnA = Split(LCase(dfn), ".") 'ext
    Select Case dfnA(UBound(dfnA)) 'type by ext
    Case "jpg", "jpeg", "png"
     send = "photo"
    Case "gif", "apng"
     send = "animation"
    Case "mp4"
     send = "video"
    Case "mp3", "m4a"
     send = "audio"
    Case "ogg"
     send = "voice"
    End Select
   End If
  End Select
  If UBound(filename) Then 'group of media files
   If Len(dfn) Then
    media = "attach://file" & i
    d.Add "file" & i, filename(i)
   Else
    media = filename(i)
   End If
   If i = 0 And Len(text) > 0 Then
    medias.Add ToD("type", send, "media", media, "caption", text)
   Else
    medias.Add ToD("type", send, "media", media)
   End If
  Else 'media file
   If Len(text) Then d.Add "caption", text
   If Len(filename(i)) Then d.Add send, filename(i)
  End If
  If d.Count > 1 Then
   If sendChatAction Then 'limits send per second
    Select Case send
    Case "photo", "voice", "video", "video_note", "document"
     tmBot token, "sendChatAction", ToD("chat_id", chat_id, "action", "upload_" & send)
     sendChatAction = False
    End Select
   End If
   If UBound(filename) Then 'group of media files
   Else
    tmBotSend = tmBotForm(token, "send" & StrConv(Replace(send, "_n", "N"), vbProperCase), d)
    Exit Function
   End If
  End If
  If (i + 1) Mod 10 = 0 Then
   d.Add "media", ConvertToJson(medias)
   tmBotSend = tmBotForm(token, "sendMediaGroup", d)
   d.RemoveAll
   d.Add "chat_id", chat_id
   Set medias = Nothing
   Set medias = New Collection
  End If
 Next
 'use module JsonConverter from https://github.com/VBA-tools/VBA-JSON
 If medias.Count Then
  d.Add "media", ConvertToJson(medias)
  tmBotSend = tmBotForm(token, "sendMediaGroup", d)
 End If
End Function

Private Function bond(Optional pref As String = vbCrLf & "--", Optional suff As String = vbCrLf, Optional BOUNDARY As String = "--OYWFRYGNCYQAOCCT44655,4239930556") As String
 'BOUNDARY string
 bond = pref & BOUNDARY & suff
End Function

Private Function form(ByVal name As String, Optional ByVal filename As String = "", Optional ByVal ct = "") As String
 'form-data string
 form = "Content-Disposition: form-data; name=""" & name & """"
 If Len(filename) Then form = form & "; filename=""" & filename & """"
 If Len(ct) Then form = form & vbCrLf & "Content-Type: " & ct
 form = form & vbCrLf & vbCrLf
End Function

Function stringToBytes(str As String) As Variant
 #If Deb Then
  Debug.Print str;
 #End If
 #If Obj Then
  With CreateObject("ADODB.Stream")
 #Else
  With New ADODB.Stream
 #End If
  .Mode = adModeReadWrite
  .Type = adTypeText
  .Charset = "UTF-8"
  .Open
  .WriteText str
  .Position = 0
  .Type = adTypeBinary
  .Position = 3 'skip BOM
  stringToBytes = .read
  .Close
 End With
End Function

Function fileToBytes(filename As String) As Variant
#If Deb Then
 Debug.Print "<" & filename & ">";
#End If
#If Obj Then
 With CreateObject("ADODB.Stream")
#Else
 With New ADODB.Stream
#End If
  .Mode = adModeReadWrite
  .Type = adTypeBinary
  .Open
  .LoadFromFile filename
  .Position = 0
  fileToBytes = .read
  .Close
 End With
End Function

Function bodyToBytes(send As String, files As Collection, bondS As String) As Variant
#If Obj Then
 With CreateObject("ADODB.Stream")
#Else
 With New ADODB.Stream
#End If
  .Mode = adModeReadWrite
  .Type = adTypeBinary
  .Open
  .write stringToBytes(send)
  For Each strA In files
   .write stringToBytes((strA(0)))
   .write fileToBytes((strA(1)))
  Next
  .write stringToBytes(bondS)
  .Position = 0
  #If Deb = 2 Then
   .SaveToFile ThisWorkbook.path & "\" & logFile, adSaveCreateOverWrite
  #End If
  bodyToBytes = .read
  .Close
 End With
End Function
#If Obj Then
Function tmBotForm(token As String, verb As String, param As Object) As String
#Else
Function tmBotForm(token As String, verb As String, param As Dictionary) As String
#End If
 Dim send As String
 Dim files As New Collection
 For Each k In param.Keys
  dfn = ""
  On Error Resume Next
  If Len(param(k)) Then dfn = Dir(param(k))
  On Error GoTo 0
  If Len(dfn) Then
   files.Add Array(bond() & form(k, dfn), param(k))
  Else
   send = send & bond() & form(k) & param(k)
  End If
 Next
#If Obj Then
 With CreateObject("MSXML2.XMLHTTP")
#Else
 With New MSXML2.XMLHTTP60
#End If
  .Open "POST", api & token & "/" & verb, False
  .setRequestHeader "Content-Type", "multipart/form-data; boundary=" & bond("", "")
  #If Deb Then
   Debug.Print "POST " & api & "<token>/" & verb
   Debug.Print "Content-Type: multipart/form-data; boundary=" & bond("", "") & String(2, vbCrLf)
   T0 = Timer
   .send bodyToBytes(send, files, bond(suff:="--"))
   'use module JsonConverter from https://github.com/VBA-tools/VBA-JSON
   Debug.Print ConvertToJson(ParseJson(.responseText), Whitespace:=1), Timer - T0
   WaitSec 'limits send per second
  #Else
  .send bodyToBytes(send, files, bond(suff:="--"))
  #End If
  tmBotForm = .responseText
 End With
End Function
#If Obj Then
Function tmBot(token As String, verb As String, Optional param As Object) As String
#Else
Function tmBot(token As String, verb As String, Optional param As Dictionary) As String
#End If
 send = api & token & "/" & verb
 If param Is Nothing Then
 Else
  sep = "?"
  For Each k In param.Keys
   send = send & sep & k & "=" & WorksheetFunction.EncodeURL(param(k))
   sep = "&"
  Next
 End If
#If Obj Then
 With CreateObject("MSXML2.XMLHTTP")
#Else
 With New MSXML2.XMLHTTP60
#End If
  .Open "POST", send, False
  #If Deb Then
   Debug.Print "POST " & Replace(send, token, "<token>")
   T0 = Timer
  .send
   'use module JsonConverter from https://github.com/VBA-tools/VBA-JSON
   Debug.Print ConvertToJson(ParseJson(.responseText), Whitespace:=1), Timer - T0
   WaitSec 'limits send per second
  #Else
  .send
  #End If
  tmBot = .responseText
 End With
End Function
#If Obj Then
Function ToD(ParamArray param()) As Object
 Set ToD = CreateObject("Scripting.Dictionary")
 'Module JsonConverter from https://github.com/VBA-tools/VBA-JSON used "New Dictionary" and "As Dictionary"
 'then add to your project class Dictionary from https://github.com/timhall/VBA-Dictionary
 'or set a reference to Microsoft Scripting Runtime
 'I used SeleniumBasic from https://github.com/florentbr/SeleniumBasic with his Dictionary and catch the bug
#Else
Function ToD(ParamArray param()) As Dictionary
 Set ToD = New Dictionary
#End If
 For i = 0 To UBound(param) Step 2
  If i + 1 <= UBound(param) Then v = param(i + 1)
  ToD.Add param(i), v
 Next
End Function

Sub WaitSec(Optional sec As Single = 1)
 T0 = Timer
 Do
  DoEvents
 Loop Until Timer - T0 >= sec
End Sub

Sub test()
 Stop
 Debug.Assert Not ParseJson(tmBot(token, "getMe"))("ok")
 'message
 'use module JsonConverter from https://github.com/VBA-tools/VBA-JSON
 Set FirstMessage = ParseJson(tmBotSend(token, chat_id, "Мы начинаем КВН"))
 
 tmBotSend token, chat_id, "Папа", ToD()
 'https://core.telegram.org/bots/api#sendmessage
 tmBot token, "sendMessage", ToD("chat_id", chat_id, "text", "Мама" & space(4096 - 8) & "Папа")
 tmBotForm token, "sendMessage", ToD("chat_id", chat_id, "text", "Мама" & space(4095 - 8) & "Папа")
 
 'photo
 tmBotSend token, chat_id, "фотка по файл ид", ToD("AgACAgIAAxkDAANIY90VxfyqwbbEP7xy9MacV5VwcTAAAp_EMRtlgOlK8gV2JnFsXYcBAAMCAAN3AAMuBA", "photo")
 tmBotSend token, chat_id, "фотка по УРЛ", ToD("https://vremya-ne-zhdet.ru/wp-content/uploads/2020/04/picture174.png", "photo")
 'https://core.telegram.org/bots/api#sendphoto
 tmBot token, "sendPhoto", ToD("chat_id", chat_id, "caption", "фотка по файл ид", "photo", "AgACAgIAAxkDAANIY90VxfyqwbbEP7xy9MacV5VwcTAAAp_EMRtlgOlK8gV2JnFsXYcBAAMCAAN3AAMuBA")
 
 'attach photo
 tmBotSend token, chat_id, "вложенная фотка", ToD("s:\01.jpg")
 'https://core.telegram.org/bots/api#sending-files
 tmBotForm token, "sendPhoto", ToD("chat_id", chat_id, "caption", "вложенная фотка", "photo", "s:\01.jpg")
 
 'attach photo as document
 tmBotSend token, chat_id, "вложенная фотка как файл", ToD("s:\01.jpg", "document")
 'https://core.telegram.org/bots/api#senddocument
 tmBotForm token, "sendDocument", ToD("chat_id", chat_id, "caption", "вложенная фотка как файл", "document", "s:\01.jpg")
 
 'attach video as animation
 tmBotSend token, chat_id, "вложенное видео как анимация", ToD("s:\abaku.mp4", "animation")
 'https://core.telegram.org/bots/api#sendanimation
 tmBotForm token, "sendAnimation", ToD("chat_id", chat_id, "animation", "s:\abaku.mp4", "caption", "вложенное видео как анимация")
 
 'photos
 tmBotSend token, chat_id, "фотки по файл ид", ToD("AgACAgIAAxkDAANIY90VxfyqwbbEP7xy9MacV5VwcTAAAp_EMRtlgOlK8gV2JnFsXYcBAAMCAAN3AAMuBA", "photo", "AgACAgIAAxkDAANiY-HtiTrOf1yGJcU3_-9H2rwDLdEAAlXFMRuxTwlLqAge0lEC0wkBAAMCAAN5AAMuBA", "photo")
 'photos raw
 tmBot token, "sendMediaGroup", ToD("chat_id", chat_id, "media", ConvertToJson(Array(ToD("caption", "фотки по файл ид", "type", "photo", "media", "AgACAgIAAxkDAANIY90VxfyqwbbEP7xy9MacV5VwcTAAAp_EMRtlgOlK8gV2JnFsXYcBAAMCAAN3AAMuBA"), ToD("type", "photo", "media", "AgACAgIAAxkDAANiY-HtiTrOf1yGJcU3_-9H2rwDLdEAAlXFMRuxTwlLqAge0lEC0wkBAAMCAAN5AAMuBA"))))
 
 'attach photos
 tmBotSend token, chat_id, "вложенные фотки", ToD("s:\01.jpg", "", "s:\02.jpg", "")
 'https://core.telegram.org/bots/api#sendmediagroup
 tmBotForm token, "sendMediaGroup", ToD("chat_id", chat_id, "media", "[{""caption"":""вложенные фотки"",""type"":""photo"",""media"":""attach://01.jpg""},{""type"":""photo"",""media"":""attach://02.jpg""}]", "01.jpg", "s:\01.jpg", "02.jpg", "s:\02.jpg")
 
 'attach documents
 tmBotSend token, chat_id, "вложенные файлы", ToD("s:\01.jpg", "document", "s:\02.jpg", "document")
 'attach documents raw
 tmBotForm token, "sendMediaGroup", ToD("chat_id", chat_id, "media", ConvertToJson(Array(ToD("caption", "вложенные файлы", "type", "document", "media", "attach://p1"), ToD("type", "document", "media", "attach://p2"))), "p1", "s:\01.jpg", "p2", "s:\02.jpg")
 
 'attach photo video
 tmBotSend token, chat_id, "фотка и видео", ToD("s:\01.jpg", "", "s:\abaku.mp4", "")
 'attach photo video raw
 tmBotForm token, "sendMediaGroup", ToD("chat_id", chat_id, "media", ConvertToJson(Array(ToD("caption", "фотка и видео", "type", "photo", "media", "attach://p"), ToD("type", "video", "media", "attach://v"))), "p", "s:\01.jpg", "v", "s:\abaku.mp4")
 
 'attach 11 photo and video
 tmBotSend token, chat_id, "11 фоток и видео", ToD("s:\01.jpg", "", "s:\02.jpg", "", "s:\04.jpg", "", "s:\05.jpg", "", "s:\07.jpg", "", "s:\08.jpg", "", "s:\09.jpg", "", "s:\11.jpg", "", "s:\12.jpg", "", "s:\13.jpg", "", "s:\14.jpg", "", "s:\abaku.mp4", "")
 
 Set lastMessage = ParseJson(tmBotSend(token, chat_id, "Расчёт окончен"))
 Stop
 If Not FirstMessage("ok") Then Exit Sub
 If Not lastMessage("ok") Then Exit Sub
 If 1 Then
  first = FirstMessage("result")("message_id")
  Last = lastMessage("result")("message_id")
 Else
  first = 270
  Last = 212
 End If
 For i = first To Last 'https://core.telegram.org/bots/api#deletemessage
  Set deleteMessage = ParseJson(tmBot(token, "deleteMessage", ToD("chat_id", chat_id, "message_id", i)))
  Debug.Print i, deleteMessage("ok")
  If Not deleteMessage("ok") Then Debug.Print deleteMessage("description")
 Next
End Sub
