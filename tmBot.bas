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
  If Not filename Like "*?:*?.?*" Then Exit Function 'file_id
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
 If Not likeFilename(filename) Then Exit Function
 On Error Resume Next
 pavd2dfn = Dir(filename)
End Function

#If Obj Then
Public Function tmBotSend(token As String, chat_id As String, Optional Text As String = vbNullString, Optional param As Object) As String
 Dim d As Object
 Dim dAll As Object
#Else
Public Function tmBotSend(token As String, chat_id As String, Optional Text As String = vbNullString, Optional param As Dictionary) As String
 Dim d As Dictionary
 Dim dAll As Dictionary
#End If
 'https://www.planetaexcel.ru/forum/index.php?PAGE_NAME=message&FID=1&TID=93149&TITLE_SEO=93149-kak-sdelat-otpravku-v-telegram-iz-makrosa-vba-excel&MID=1193376#message1193376
 'use ToD(key, value, key2, value2, ...) to setup param
 Set dAll = ToD("chat_id", chat_id)
 Dim medias As New Collection
 Dim files As New Collection
 Dim send As String
 Dim dfn As String
 Dim caption As Boolean: caption = Len(Text) 'only once
 Dim sendChatAction As Boolean: sendChatAction = True 'only once
 Dim doSend As Boolean
 Dim gi As Integer 'group index
 If param Is Nothing Then
  If caption Then
   dAll.Add "text", Text
   tmBotSend = tmBotURL(token, "sendMessage", dAll)
  End If
  Exit Function
 End If
 attach = 0
 For Each k In param.keys
  dfn = pavd2dfn(param(k), k, send) 'filename without path case exist
  If Len(send) > 0 Then
   If Len(dfn) Then attach = attach + 1
   files.Add Array(k, send, dfn) 'path|url|file_id type filename
  Else
   dAll.Add k, param(k)
  End If
 Next
 If files.Count = 0 Then
  If caption Then
   dAll.Add "text", Text
   tmBotSend = tmBotForm(token, "sendMessage", dAll)
  End If
  Exit Function
 End If
 Set d = ToD()
 For Each k In dAll.keys
  d.Add k, dAll(k)
 Next
 gi = 0
 For i = 1 To files.Count
  filename = files(i)(0)
  send = files(i)(1)
  dfn = files(i)(2)
  If attach > 0 And sendChatAction Then
   Select Case send
   Case "animation", "audio"
   Case Else
    tmBotURL token, "sendChatAction", ToD("chat_id", chat_id, "action", "upload_" & send)
    sendChatAction = False
   End Select
  End If
  If files.Count = 1 Then
   If likeFilename(filename) And Len(dfn) = 0 Then Exit Function
   If caption Then d.Add "caption", Text
   d.Add send, filename
   tmBotSend = tmBotForm(token, "send" & Replace(send, "_n", "N"), d)
   Exit Function
  End If
  'group of media files
  If Len(dfn) Then
   media = "attach://file" & i
   d.Add "file" & i, filename
  Else
   If likeFilename(filename) Then
    media = vbNullString
   Else
    media = filename
   End If
  End If
  If Len(media) Then
   gi = gi + 1
   If caption Then
    medias.Add ToD("type", send, "media", media, "caption", Text)
    caption = False
   Else
    medias.Add ToD("type", send, "media", media)
   End If
  End If
  If gi Mod 10 = 0 Then
   tmBotSend = sendMediaGroup(medias, d)
   d.RemoveAll
   For Each k In dAll.keys
    d.Add k, dAll(k)
   Next
   Set medias = New Collection
  End If
 Next
 tmBotSend = sendMediaGroup(medias, d)
End Function

#If Obj Then
Function sendMediaGroup(medias As Collection, d As Object) As String
#Else
Function sendMediaGroup(medias As Collection, d As Dictionary) As String
#End If
 If medias.Count = 0 Then Exit Function
 'use module JsonConverter from https://github.com/VBA-tools/VBA-JSON
 d.Add "media", ConvertToJson(medias)
 sendMediaGroup = tmBotForm(token, "sendMediaGroup", d)
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
Function likeFilename(filename) As Boolean
 likeFilename = filename Like "[A-z]:\?*"
End Function
#If Obj Then
Function tmBotForm(token As String, verb As String, param As Object) As String
#Else
Function tmBotForm(token As String, verb As String, param As Dictionary) As String
#End If
 Dim multipart As String
 Dim files As New Collection
 Dim dfn As String
 Dim send As String
 For Each k In param.keys
  If VarType(k) = vbString Then
   dfn = pavd2dfn(k, param(k), send)
   If Len(dfn) Then
    files.Add Array(bond() & form(k, dfn), param(k))
   Else
    If Len(send) > 0 And likeFilename(param(k)) Then
     Exit Function
    Else
     multipart = multipart & bond() & form(k) & param(k)
    End If
   End If
  End If
 Next
 If files.Count = 0 Then 'URL query string
  tmBotForm = tmBotURL(token, verb, param)
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
   Debug.Print
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
Function tmBotURL(token As String, verb As String, Optional param As Object) As String
#Else
Function tmBotURL(token As String, verb As String, Optional param As Dictionary) As String
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
  tmBotURL = .responseText
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
 json = tmBotURL(token, "getMe")
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
 json = tmBotSend(token, chat_id, "01 ???? ???????????????? ??????")
 firstMessage = ParseJsonPart(json, "message_id")
 Debug.Print ParseJsonPart(json, "text")
 Debug.Print timestamp2date(ParseJsonPart(json, "date"))
 Stop
 Debug.Print ParseJsonPart(json, "from", "username")
 Debug.Print ParseJsonPart(json, "chat", "username")
 Stop
 tmBotSend token, chat_id, "02 ????????", ToD()
 'https://core.telegram.org/bots/api#sendmessage
 tmBotURL token, "sendMessage", ToD("chat_id", chat_id, "text", "03" & space(4096 - 6) & "????????")
 tmBotForm token, "sendMessage", ToD("chat_id", chat_id, "text", "04" & space(4096 - 6 - 1) & "????????")
 
 'photo
 tmBotSend token, chat_id, "05 ?????????? ???? ???????? ????", ToD("AgACAgIAAxkDAANIY90VxfyqwbbEP7xy9MacV5VwcTAAAp_EMRtlgOlK8gV2JnFsXYcBAAMCAAN3AAMuBA", "photo")
 tmBotSend token, chat_id, "06 ?????????? ???? ??????", ToD("https://vremya-ne-zhdet.ru/wp-content/uploads/2020/04/picture174.png")
 'https://core.telegram.org/bots/api#sendphoto
 tmBotURL token, "sendPhoto", ToD("chat_id", chat_id, "caption", "07 ?????????? ???? ???????? ????", "photo", "AgACAgIAAxkDAANIY90VxfyqwbbEP7xy9MacV5VwcTAAAp_EMRtlgOlK8gV2JnFsXYcBAAMCAAN3AAMuBA")
 
 'attach photo
 tmBotSend token, chat_id, "08 ?????????????????? ??????????", ToD("s:\01.jpg")
 'https://core.telegram.org/bots/api#sending-files
 tmBotForm token, "sendPhoto", ToD("chat_id", chat_id, "caption", "09 ?????????????????? ??????????", "photo", "s:\01.jpg")
 
 'attach photo as document
 tmBotSend token, chat_id, "10 ?????????????????? ?????????? ?????? ????????", ToD("s:\01.jpg", "document")
 'https://core.telegram.org/bots/api#senddocument
 tmBotForm token, "sendDocument", ToD("chat_id", chat_id, "caption", "11 ?????????????????? ?????????? ?????? ????????", "document", "s:\01.jpg")
 
 'try attach not exist photo
 tmBotSend token, chat_id, "12 ?????????????? ?????????????? ?????????????????????????? ??????????", ToD("s:\00.jpg")
 tmBotForm token, "sendPhoto", ToD("chat_id", chat_id, "caption", "13 ?????????????? ?????????????? ?????????????????????????? ??????????", "photo", "s:\00.jpg")
 Stop
 'attach video as animation
 tmBotSend token, chat_id, "14 ?????????????????? ?????????? ?????? ????????????????", ToD("s:\abaku.mp4", "animation")
 'https://core.telegram.org/bots/api#sendanimation
 tmBotForm token, "sendAnimation", ToD("chat_id", chat_id, "animation", "s:\abaku.mp4", "caption", "15 ?????????????????? ?????????? ?????? ????????????????")
 
 'photos
 tmBotSend token, chat_id, "16 ?????????? ???? ???????? ????", ToD("AgACAgIAAxkDAANIY90VxfyqwbbEP7xy9MacV5VwcTAAAp_EMRtlgOlK8gV2JnFsXYcBAAMCAAN3AAMuBA", "photo", "AgACAgIAAxkDAANiY-HtiTrOf1yGJcU3_-9H2rwDLdEAAlXFMRuxTwlLqAge0lEC0wkBAAMCAAN5AAMuBA", "photo")
 'photos raw
 tmBotURL token, "sendMediaGroup", ToD("chat_id", chat_id, "media", ConvertToJson(Array(ToD("caption", "17 ?????????? ???? ???????? ????", "type", "photo", "media", "AgACAgIAAxkDAANIY90VxfyqwbbEP7xy9MacV5VwcTAAAp_EMRtlgOlK8gV2JnFsXYcBAAMCAAN3AAMuBA"), ToD("type", "photo", "media", "AgACAgIAAxkDAANiY-HtiTrOf1yGJcU3_-9H2rwDLdEAAlXFMRuxTwlLqAge0lEC0wkBAAMCAAN5AAMuBA"))))
 
 'attach photos
 tmBotSend token, chat_id, "18 ?????????????????? ??????????", ToD("s:\01.jpg", "", "s:\02.jpg")
 'https://core.telegram.org/bots/api#sendmediagroup
 tmBotForm token, "sendMediaGroup", ToD("chat_id", chat_id, "media", "[{""caption"":""19 ?????????????????? ??????????"",""type"":""photo"",""media"":""attach://01.jpg""},{""type"":""photo"",""media"":""attach://02.jpg""}]", "01.jpg", "s:\01.jpg", "02.jpg", "s:\02.jpg")
 
 'attach documents
 tmBotSend token, chat_id, "20 ?????????????????? ??????????", ToD("s:\01.jpg", "document", "s:\02.jpg", "document")
 'attach documents raw
 tmBotForm token, "sendMediaGroup", ToD("chat_id", chat_id, "media", ConvertToJson(Array(ToD("caption", "21 ?????????????????? ??????????", "type", "document", "media", "attach://p1"), ToD("type", "document", "media", "attach://p2"))), "p1", "s:\01.jpg", "p2", "s:\02.jpg")
 
 'attach photo video
 tmBotSend token, chat_id, "22 ?????????? ?? ??????????", ToD("s:\01.jpg", "", "s:\abaku.mp4", "")
 'attach photo video raw
 tmBotForm token, "sendMediaGroup", ToD("chat_id", chat_id, "media", ConvertToJson(Array(ToD("caption", "23 ?????????? ?? ??????????", "type", "photo", "media", "attach://p"), ToD("type", "video", "media", "attach://v"))), "p", "s:\01.jpg", "v", "s:\abaku.mp4")
 
 'try attach photo and unexist photo
 tmBotSend token, chat_id, "24 ?????????????? ?????????????? ?????????? ?? ?????????????????????????? ??????????", ToD("s:\01.jpg", "", "s:\00.jpg")
 'attach 11 photo and video
 tmBotSend token, chat_id, "25 ?????????? ?? 11 ??????????", ToD("s:\abaku.mp4", "", "s:\01.jpg", "", "s:\02.jpg", "", "s:\04.jpg", "", "s:\05.jpg", "", "s:\07.jpg", "", "s:\08.jpg", "", "s:\09.jpg", "", "s:\11.jpg", "", "s:\12.jpg", "", "s:\13.jpg", "", "s:\14.jpg", "", "s:\00.jpg")
#If JSON_Parser_by_Daniel_Ferry Then
 lastMessage = Val(json2dic(tmBotSend(token, chat_id, "???????????? ??????????????"))("obj.result.message_id"))
End If
 lastMessage = ParseJsonPart(tmBotSend(token, chat_id, "???????????? ??????????????"), "message_id")
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
  Set deleteMessage = ParseJSON(tmBotURL(token, "deleteMessage", ToD("chat_id", chat_id, "message_id", i)))
  Debug.Print i, deleteMessage("ok")
  If Not deleteMessage("ok") Then Debug.Print deleteMessage("description")
 Next
End Sub
