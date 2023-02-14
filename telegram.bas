#Const JSON_Parser_by_Daniel_Ferry = False
#Const Obj = True
#If Obj Then
 Const adTypeBinary = 1
 Const adTypeText = 2
 Const adModeReadWrite = 3
 Const adSaveCreateOverWrite = 2
#Else
'set reference to:
'Sripting.Dictionary
'ADODB.Stream
'MSXML2.XMLHTTP
#End If
Const api = "https://api.telegram.org/bot"
#Const Deb = 1 'limits send per second and do debug.print
Const logFile = "body.txt" 'if Deb=2 then save body to logfile

Function pavd2dfn(ByVal pavd As String, ByVal filename As String, Optional ByRef send As String) As String
 send = vbNullString
 If Len(filename) = 0 Then Exit Function
 Select Case LCase(pavd)
 Case "photo", "animation", "audio", "voice", "video", "video_note", "document"
  send = LCase(pavd)
 Case Else
  If Not filename Like "*?:*?.?*" Then Exit Function 'not file and not url
  a = Split(filename, ".")
  Select Case a(UBound(a))  'type by ext
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
 End Select
 If Len(send) = 0 Then Exit Function
 If Not filename Like "[A-z]:\?*" Then Exit Function 'not file
 On Error Resume Next
 pavd2dfn = Dir(filename)
End Function

#If Obj Then
Public Function tmBotSend(token As String, chat_id As String, Optional Text As String = vbNullString, Optional param As Object) As String
 Dim d As Object
#Else
Public Function tmBotSend(token As String, chat_id As String, Optional Text As String = vbNullString, Optional param As Dictionary) As String
 Dim d As Dictionary
#End If
 'https://www.planetaexcel.ru/forum/index.php?PAGE_NAME=message&FID=1&TID=93149&TITLE_SEO=93149-kak-sdelat-otpravku-v-telegram-iz-makrosa-vba-excel&MID=1193376#message1193376
 'use ToD(key, value, key2, value2, ...) to setup param
 Set d = ToD("chat_id", chat_id)
 Set medias = New Collection
 Dim send As String
 Dim dfn As String
 If param Is Nothing Then
  If Len(Text) Then
   d.Add "text", Text
   tmBotSend = tmBot(token, "sendMessage", d)
  End If
  Exit Function
 Else
  If param.Count = 0 Then
   If Len(Text) Then
    d.Add "text", Text
    tmBotSend = tmBotForm(token, "sendMessage", d)
   End If
   Exit Function
  End If
  filename = param.keys
  pavd = param.Items  'pavd as "photo", "animation", "audio", "voice", "video", "video_note", "document"
 End If
 sendChatAction = True 'only one sendChatAction
 For i = 0 To UBound(filename)
  dfn = pavd2dfn(pavd(i), filename(i), send) 'filename w/o path case exist
  doSend = Len(send) > 0
  If doSend Then
   If UBound(filename) Then 'group of media files
    If Len(dfn) Then
     media = "attach://file" & i
     d.Add "file" & i, filename(i)
    Else
     If filename(i) Like "[A-z]:\?*" Then 'file not exist
      doSend = False
     Else
      media = filename(i)
     End If
    End If
    If doSend Then
     If i = 0 And Len(Text) > 0 Then
      medias.Add ToD("type", send, "media", media, "caption", Text)
     Else
      medias.Add ToD("type", send, "media", media)
     End If
    End If
   Else 'media file
    If Len(Text) Then d.Add "caption", Text
    d.Add send, filename(i)
    tmBotSend = tmBotForm(token, "send" & StrConv(Replace(send, "_n", "N"), vbProperCase), d)
    Exit Function
   End If 'UBound(filename)
   If d.Count > 1 Then 'chat_id
    If sendChatAction Then 'limits send per second
     Select Case send
     Case "photo", "voice", "video", "video_note", "document"
      tmBot token, "sendChatAction", ToD("chat_id", chat_id, "action", "upload_" & send)
      sendChatAction = False
     End Select
    End If
   End If 'd.Count > 1
  End If 'doSend
  If (i + 1) Mod 10 = 0 Then
   'use module JsonConverter from https://github.com/VBA-tools/VBA-JSON
   d.Add "media", ConvertToJson(medias)
   tmBotSend = tmBotForm(token, "sendMediaGroup", d)
   d.RemoveAll
   d.Add "chat_id", chat_id
   Set medias = Nothing
   Set medias = New Collection
  End If
 Next
 If medias.Count Then
  'use module JsonConverter from https://github.com/VBA-tools/VBA-JSON
  d.Add "media", ConvertToJson(medias)
  tmBotSend = tmBotForm(token, "sendMediaGroup", d)
 End If
End Function

Private Function bond(Optional pref As String = vbCrLf & "--", Optional suff As String = vbCrLf, Optional BOUNDARY As String = "--OYWFRYGNCYQAOCCT44655,4239930556") As String
 'BOUNDARY string
 bond = pref & BOUNDARY & suff
End Function

Private Function form(ByVal name As String, Optional ByVal filename As String = vbNullString) As String
 'form-data string
 form = "Content-Disposition: form-data; name=""" & name & """"
 If Len(filename) Then form = form & "; filename=""" & filename & """"
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
 Dim multipart As String
 Dim files As New Collection
 Dim doSend As Boolean: doSend = True
 Dim dfn As String
 For Each k In param.keys
  dfn = pavd2dfn(k, param(k))
  If Len(dfn) Then
   files.Add Array(bond() & form(k, dfn), param(k))
  Else
   If param(k) Like "[A-z]:\?*" Then 'file not exist
    doSend = False
   Else
    multipart = multipart & bond() & form(k) & param(k)
   End If
  End If
 Next
 If Not doSend Then Exit Function
 If files.Count = 0 Then
  tmBotForm = tmBot(token, verb, param)
  Exit Function
 End If
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
   .send bodyToBytes(multipart, files, bond(suff:="--"))
   'use module JsonConverter from https://github.com/VBA-tools/VBA-JSON
   Debug.Print ConvertToJson(ParseJSON(.responseText), Whitespace:=1), Timer - T0
   WaitSec 'limits send per second
  #Else
  .send bodyToBytes(multipart, files, bond(suff:="--"))
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
  For Each k In param.keys
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
   Debug.Print ConvertToJson(ParseJSON(.responseText), Whitespace:=1), Timer - T0
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
#Else
Function ToD(ParamArray param()) As Dictionary
 Set ToD = New Scripting.Dictionary
#End If
 On Error Resume Next
 For i = 0 To UBound(param) Step 2
  v = vbNullString
  v = param(i + 1)
  ToD.Add param(i), v
 Next
End Function

Sub WaitSec(Optional sec As Single = 1)
 T0 = Timer
 Do
  DoEvents
 Loop Until Timer - T0 >= sec
End Sub

Function timestamp2date(ByVal value) As Date
 timestamp2date = ParseUtc(CDate(value / 60 / 60 / 24) + "1/1/1970")
End Function

Sub test()
 Stop
 Dim json As String
 json = tmBot(token, "getMe")
#If JSON_Parser_by_Daniel_Ferry Then
 Set dic = json2dic(json)
 Debug.Print ListPaths(dic)
 Debug.Assert dic("obj.ok") = "true"
 For Each v In GetFilteredValues(dic, "obj.result*")
  Debug.Print v
 Next
#End If
 'like ParseJson(json)("ok") from https://github.com/VBA-tools/VBA-JSON but without parse all json
 Debug.Assert ParseJsonPart(json, "ok") 'https://github.com/abakum/VBA-JSON/blob/master/JsonConverter.bas
 'get first value by key "id" - dirty hack
 Debug.Print ParseJsonPart(json, "id")
 'like ParseJson(json)("result")("id") but without parse all json
 Debug.Print ParseJsonPart(json, "result", "id")
 Stop
 'message
 json = tmBotSend(token, chat_id, "01 Мы начинаем КВН")
 firstMessage = ParseJsonPart(json, "message_id")
 Debug.Print ParseJsonPart(json, "text")
 Debug.Print timestamp2date(ParseJsonPart(json, "date"))
 Stop
 Debug.Print ParseJsonPart(json, "from", "username")
 Debug.Print ParseJsonPart(json, "chat", "username")
 Stop
 tmBotSend token, chat_id, "02 Мама", ToD()
 'https://core.telegram.org/bots/api#sendmessage
 tmBot token, "sendMessage", ToD("chat_id", chat_id, "text", "03" & space(4096 - 6) & "Мама")
 tmBotForm token, "sendMessage", ToD("chat_id", chat_id, "text", "04" & space(4096 - 6 - 1) & "Мама")
 
 'photo
 tmBotSend token, chat_id, "05 фотка по файл ид", ToD("AgACAgIAAxkDAANIY90VxfyqwbbEP7xy9MacV5VwcTAAAp_EMRtlgOlK8gV2JnFsXYcBAAMCAAN3AAMuBA", "photo")
 tmBotSend token, chat_id, "06 фотка по УРЛ", ToD("https://vremya-ne-zhdet.ru/wp-content/uploads/2020/04/picture174.png")
 'https://core.telegram.org/bots/api#sendphoto
 tmBot token, "sendPhoto", ToD("chat_id", chat_id, "caption", "07 фотка по файл ид", "photo", "AgACAgIAAxkDAANIY90VxfyqwbbEP7xy9MacV5VwcTAAAp_EMRtlgOlK8gV2JnFsXYcBAAMCAAN3AAMuBA")
 
 'attach photo
 tmBotSend token, chat_id, "08 вложенная фотка", ToD("s:\01.jpg")
 'https://core.telegram.org/bots/api#sending-files
 tmBotForm token, "sendPhoto", ToD("chat_id", chat_id, "caption", "09 вложенная фотка", "photo", "s:\01.jpg")
 
 'attach photo as document
 tmBotSend token, chat_id, "10 вложенная фотка как файл", ToD("s:\01.jpg", "document")
 'https://core.telegram.org/bots/api#senddocument
 tmBotForm token, "sendDocument", ToD("chat_id", chat_id, "caption", "11 вложенная фотка как файл", "document", "s:\01.jpg")
 
 'try attach not exist photo
 tmBotSend token, chat_id, "12 попытка послать отсутствующую фотку", ToD("s:\00.jpg")
 tmBotForm token, "sendPhoto", ToD("chat_id", chat_id, "caption", "13 попытка послать отсутствующую фотку", "photo", "s:\00.jpg")
 
 'attach video as animation
 tmBotSend token, chat_id, "14 вложенное видео как анимация", ToD("s:\abaku.mp4", "animation")
 'https://core.telegram.org/bots/api#sendanimation
 tmBotForm token, "sendAnimation", ToD("chat_id", chat_id, "animation", "s:\abaku.mp4", "caption", "15 вложенное видео как анимация")
 
 'photos
 tmBotSend token, chat_id, "16 фотки по файл ид", ToD("AgACAgIAAxkDAANIY90VxfyqwbbEP7xy9MacV5VwcTAAAp_EMRtlgOlK8gV2JnFsXYcBAAMCAAN3AAMuBA", "photo", "AgACAgIAAxkDAANiY-HtiTrOf1yGJcU3_-9H2rwDLdEAAlXFMRuxTwlLqAge0lEC0wkBAAMCAAN5AAMuBA", "photo")
 'photos raw
 tmBot token, "sendMediaGroup", ToD("chat_id", chat_id, "media", ConvertToJson(Array(ToD("caption", "17 фотки по файл ид", "type", "photo", "media", "AgACAgIAAxkDAANIY90VxfyqwbbEP7xy9MacV5VwcTAAAp_EMRtlgOlK8gV2JnFsXYcBAAMCAAN3AAMuBA"), ToD("type", "photo", "media", "AgACAgIAAxkDAANiY-HtiTrOf1yGJcU3_-9H2rwDLdEAAlXFMRuxTwlLqAge0lEC0wkBAAMCAAN5AAMuBA"))))
 
 'attach photos
 tmBotSend token, chat_id, "18 вложенные фотки", ToD("s:\01.jpg", "", "s:\02.jpg", "")
 'https://core.telegram.org/bots/api#sendmediagroup
 tmBotForm token, "sendMediaGroup", ToD("chat_id", chat_id, "media", "[{""caption"":""19 вложенные фотки"",""type"":""photo"",""media"":""attach://01.jpg""},{""type"":""photo"",""media"":""attach://02.jpg""}]", "01.jpg", "s:\01.jpg", "02.jpg", "s:\02.jpg")
 
 'attach documents
 tmBotSend token, chat_id, "20 вложенные файлы", ToD("s:\01.jpg", "document", "s:\02.jpg", "document")
 'attach documents raw
 tmBotForm token, "sendMediaGroup", ToD("chat_id", chat_id, "media", ConvertToJson(Array(ToD("caption", "21 вложенные файлы", "type", "document", "media", "attach://p1"), ToD("type", "document", "media", "attach://p2"))), "p1", "s:\01.jpg", "p2", "s:\02.jpg")
 
 'attach photo video
 tmBotSend token, chat_id, "22 фотка и видео", ToD("s:\01.jpg", "", "s:\abaku.mp4", "")
 'attach photo video raw
 tmBotForm token, "sendMediaGroup", ToD("chat_id", chat_id, "media", ConvertToJson(Array(ToD("caption", "23 фотка и видео", "type", "photo", "media", "attach://p"), ToD("type", "video", "media", "attach://v"))), "p", "s:\01.jpg", "v", "s:\abaku.mp4")
 
 'try attach photo and unexist photo
 tmBotSend token, chat_id, "24 попытка послать фотку и отсутствующую фотку", ToD("s:\01.jpg", "", "s:\00.jpg")
 'attach 11 photo and video
 tmBotSend token, chat_id, "25 видео и 11 фоток", ToD("s:\abaku.mp4", "", "s:\01.jpg", "", "s:\02.jpg", "", "s:\04.jpg", "", "s:\05.jpg", "", "s:\07.jpg", "", "s:\08.jpg", "", "s:\09.jpg", "", "s:\11.jpg", "", "s:\12.jpg", "", "s:\13.jpg", "", "s:\14.jpg")
 
 'lastMessage = Val(json2dic(tmBotSend(token, chat_id, "Расчёт окончен"))("obj.result.message_id"))
 lastMessage = ParseJsonPart(tmBotSend(token, chat_id, "Расчёт окончен"), "message_id")
 Stop
 If IsNull(firstMessage) Then Exit Sub
 If IsNull(lastMessage) Then Exit Sub
 If 1 Then
  first = firstMessage
  Last = lastMessage
 Else
  first = 270
  Last = 212
 End If
 For i = first To Last 'https://core.telegram.org/bots/api#deletemessage
  Set deleteMessage = ParseJSON(tmBot(token, "deleteMessage", ToD("chat_id", chat_id, "message_id", i)))
  Debug.Print i, deleteMessage("ok")
  If Not deleteMessage("ok") Then Debug.Print deleteMessage("description")
 Next
End Sub
