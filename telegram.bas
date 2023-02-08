Attribute VB_Name = "telegram"
Const adTypeBinary = 1
Const adTypeText = 2
Const adModeReadWrite = 3
Const adSaveCreateOverWrite = 2
Const telegram = "https://api.telegram.org/bot"
Const tmDeb = True

Public Function tmBotSend(Token As String, chat_id As String, Optional text As String = "", Optional param As Object) As String
 'https://www.planetaexcel.ru/forum/index.php?PAGE_NAME=message&FID=1&TID=93149&TITLE_SEO=93149-kak-sdelat-otpravku-v-telegram-iz-makrosa-vba-excel&MID=1193376#message1193376
 'param is Dictionary use ToD()
 Dim d As Object
 Set d = ToD()
 Set medias = New Collection
 If param Is Nothing Then
  If Len(text) Then tmBotSend = tmBot(Token, chat_id, "sendMessage", ToD("text", text))
  Exit Function
 Else
  If param.Count = 0 Then
   If Len(text) Then tmBotSend = tmBotForm(Token, chat_id, "sendMessage", ToD("text", text))
   Exit Function
  End If
  filename = param.Keys
  pavd = param.Items  'pavd as "photo", "animation", "audio", "voice", "video", "video_note" and "document" as default
 End If
 For i = 0 To IIf(UBound(filename) > 9, 9, UBound(filename))
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
  If d.Count Then
   Select Case send
   Case "photo", "voice", "video", "video_note", "document"
    tmBot Token, chat_id, "sendChatAction", ToD("action", "upload_" & send)
   End Select
   If UBound(filename) Then 'group of media
   Else
    tmBotSend = tmBotForm(Token, chat_id, "send" & StrConv(Replace(send, "_n", "N"), vbProperCase), d)
    Exit Function
   End If
  End If
 Next
 d.Add "media", ConvertToJson(medias)
 tmBotSend = tmBotForm(Token, chat_id, "sendMediaGroup", d)
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
 If tmDeb Then Debug.Print str, ;
 With CreateObject("ADODB.Stream")
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
 If tmDeb Then Debug.Print "`" & filename & "`"
 With CreateObject("ADODB.Stream")
  .Mode = adModeReadWrite
  .Type = adTypeBinary
  .Open
  .LoadFromFile filename
  .Position = 0
  fileToBytes = .read
  .Close
 End With
End Function

Function bodyToBytes(send As String, fileC As Collection, bondS As String) As Variant
 With CreateObject("ADODB.Stream")
  .Mode = adModeReadWrite
  .Type = adTypeBinary
  .Open
  .write stringToBytes(send)
  For Each arr In fileC
   .write stringToBytes(CStr(arr(0)))
   .write fileToBytes(CStr(arr(1)))
  Next
  .write stringToBytes(bondS)
  .Position = 0
  If tmDeb Then .SaveToFile "c:\ins\body.txt", adSaveCreateOverWrite
  bodyToBytes = .read
  .Close
 End With
End Function

Function tmBotForm(Token As String, chat_id As String, verb As String, param As Object) As String
 'param is Dictionary use ToD()
 Dim send As String
 Dim fileC As New Collection
 send = bond("--") & form("chat_id") & chat_id
 For Each k In param.Keys
  dfn = ""
  On Error Resume Next
  If Len(param(k)) Then dfn = Dir(param(k))
  On Error GoTo 0
  If Len(dfn) Then
   fileC.Add Array(bond() & form(k, dfn), param(k))
  Else
   send = send & bond() & form(k) & param(k)
  End If
 Next
 With CreateObject("MSXML2.XMLHTTP")
  .Open "POST", telegram & Token & "/" & verb, False
  .setRequestHeader "Content-Type", "multipart/form-data; boundary=" & bond("", "")
  .send bodyToBytes(send, fileC, bond(suff:="--"))
  tmBotForm = .responseText
  If tmDeb Then Debug.Print ConvertToJson(ParseJson(.responseText), Whitespace:=1)
 End With
End Function
Function tmBot(Token As String, chat_id As String, verb As String, param As Object) As String
 'param is Dictionary use ToD()
 send = telegram & Token & "/" & verb & "?chat_id=" & chat_id
 For Each k In param.Keys
  send = send & "&" & k & "=" & WorksheetFunction.EncodeURL(param(k))
 Next
 If tmDeb Then Debug.Print send
 With CreateObject("MSXML2.XMLHTTP")
  .Open "POST", send, False
  .send
  tmBot = .responseText
  If tmDeb Then Debug.Print ConvertToJson(ParseJson(.responseText), Whitespace:=1)
 End With
End Function
Function ToScD(ParamArray param()) As Object
 Set ToScD = CreateObject("Scripting.Dictionary")
 For i = 0 To UBound(param) Step 2
  If i + 1 <= UBound(param) Then v = param(i + 1)
  ToScD.Add param(i), v
 Next
End Function
Function ToSeD(ParamArray param()) As Object
 Set ToSeD = New Selenium.Dictionary
 For i = 0 To UBound(param) Step 2
  If i + 1 <= UBound(param) Then v = param(i + 1)
  ToSeD.Add param(i), v
 Next
End Function
Function ToD(ParamArray param()) As Object
 Set ToD = New Dictionary
 For i = 0 To UBound(param) Step 2
  If i + 1 <= UBound(param) Then v = param(i + 1)
  ToD.Add param(i), v
 Next
End Function

Sub test()
 Stop
 'message
 Set FirstMessage = ParseJson(tmBotSend(Token, chat_id, "Мы начинаем КВН"))
 tmBotSend Token, chat_id, "Папа", ToD()
 'https://core.telegram.org/bots/api#sendmessage
 tmBot Token, chat_id, "sendMessage", ToD("text", "Мама" & space(4096 - 8) & "Папа")
 tmBotForm Token, chat_id, "sendMessage", ToD("text", "Мама" & space(4095 - 8) & "Папа")
 
 'photo
 tmBotSend Token, chat_id, "фотка по файл ид", ToD("AgACAgIAAxkDAANIY90VxfyqwbbEP7xy9MacV5VwcTAAAp_EMRtlgOlK8gV2JnFsXYcBAAMCAAN3AAMuBA", "photo")
 tmBotSend Token, chat_id, "фотка по УРЛ", ToD("https://vremya-ne-zhdet.ru/wp-content/uploads/2020/04/picture174.png", "photo")
 'https://core.telegram.org/bots/api#sendphoto
 tmBot Token, chat_id, "sendPhoto", ToD("caption", "фотка по файл ид", "photo", "AgACAgIAAxkDAANIY90VxfyqwbbEP7xy9MacV5VwcTAAAp_EMRtlgOlK8gV2JnFsXYcBAAMCAAN3AAMuBA")
 
 'attach photo
 tmBotSend Token, chat_id, "вложенная фотка", ToD("s:\01.jpg")
 'https://core.telegram.org/bots/api#sending-files
 tmBotForm Token, chat_id, "sendPhoto", ToD("caption", "вложенная фотка", "photo", "s:\01.jpg")
 
 'attach photo as document
 tmBotSend Token, chat_id, "вложенная фотка как файл", ToD("s:\01.jpg", "document")
 'https://core.telegram.org/bots/api#senddocument
 tmBotForm Token, chat_id, "sendDocument", ToD("caption", "вложенная фотка как файл", "document", "s:\01.jpg")
 
 'attach video as animation
 tmBotSend Token, chat_id, "вложенное видео как анимация", ToD("s:\abaku.mp4", "animation")
 'https://core.telegram.org/bots/api#sendanimation
 tmBotForm Token, chat_id, "sendAnimation", ToD("animation", "s:\abaku.mp4", "caption", "вложенное видео как анимация")
 
 'photos
 tmBotSend Token, chat_id, "фотки по файл ид", ToD("AgACAgIAAxkDAANIY90VxfyqwbbEP7xy9MacV5VwcTAAAp_EMRtlgOlK8gV2JnFsXYcBAAMCAAN3AAMuBA", "photo", "AgACAgIAAxkDAANiY-HtiTrOf1yGJcU3_-9H2rwDLdEAAlXFMRuxTwlLqAge0lEC0wkBAAMCAAN5AAMuBA", "photo")
 'photos raw
 tmBot Token, chat_id, "sendMediaGroup", ToD("media", ConvertToJson(Array(ToD("caption", "фотки по файл ид", "type", "photo", "media", "AgACAgIAAxkDAANIY90VxfyqwbbEP7xy9MacV5VwcTAAAp_EMRtlgOlK8gV2JnFsXYcBAAMCAAN3AAMuBA"), ToD("type", "photo", "media", "AgACAgIAAxkDAANiY-HtiTrOf1yGJcU3_-9H2rwDLdEAAlXFMRuxTwlLqAge0lEC0wkBAAMCAAN5AAMuBA"))))
 
 'attach photos
 tmBotSend Token, chat_id, "вложенные фотки", ToD("s:\01.jpg", "", "s:\02.jpg", "")
 'https://core.telegram.org/bots/api#sendmediagroup
 tmBotForm Token, chat_id, "sendMediaGroup", ToD("media", "[{""caption"":""вложенные фотки"",""type"":""photo"",""media"":""attach://01.jpg""},{""type"":""photo"",""media"":""attach://02.jpg""}]", "01.jpg", "s:\01.jpg", "02.jpg", "s:\02.jpg")
 
 'attach documents
 tmBotSend Token, chat_id, "вложенные файлы", ToD("s:\01.jpg", "document", "s:\02.jpg", "document")
 'attach documents raw
 tmBotForm Token, chat_id, "sendMediaGroup", ToD("media", ConvertToJson(Array(ToD("caption", "вложенные файлы", "type", "document", "media", "attach://p1"), ToD("type", "document", "media", "attach://p2"))), "p1", "s:\01.jpg", "p2", "s:\02.jpg")
 
 'attach photo video
 tmBotSend Token, chat_id, "фотка и видео", ToD("s:\01.jpg", "", "s:\abaku.mp4", "")
 'attach photo video raw
 tmBotForm Token, chat_id, "sendMediaGroup", ToD("media", ConvertToJson(Array(ToD("caption", "фотка и видео", "type", "photo", "media", "attach://p"), ToD("type", "video", "media", "attach://v"))), "p", "s:\01.jpg", "v", "s:\abaku.mp4")
 Set lastMessage = ParseJson(tmBotSend(Token, chat_id, "Расчёт окончен"))
 Stop
 If Not FirstMessage("ok") Then Exit Sub
 If Not lastMessage("ok") Then Exit Sub
 If 1 Then
  First = FirstMessage("result")("message_id")
  Last = lastMessage("result")("message_id")
 Else
  First = 210
  Last = 212
 End If
 For i = First To Last 'https://core.telegram.org/bots/api#deletemessage
  Set deleteMessage = ParseJson(tmBot(Token, chat_id, "deleteMessage", ToD("message_id", i)))
  Debug.Print i, deleteMessage("ok")
  If Not deleteMessage("ok") Then Debug.Print deleteMessage("description")
 Next
End Sub


