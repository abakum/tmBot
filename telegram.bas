'Const adTypeBinary = 1
'Const adTypeText = 2
'Const adModeReadWrite = 3
'Const adSaveCreateOverWrite = 2
Const api = "https://api.telegram.org/bot"
Const tmDeb = 1 'limits send per second and do debug.print
Const logFile = "c:\ins\body.txt" 'if tmDep=2 then save body to logfile

Public Function tmBotSend(token As String, chat_id As String, Optional text As String = "", Optional param As Dictionary) As String
 'https://www.planetaexcel.ru/forum/index.php?PAGE_NAME=message&FID=1&TID=93149&TITLE_SEO=93149-kak-sdelat-otpravku-v-telegram-iz-makrosa-vba-excel&MID=1193376#message1193376
 'use ToD() to setup param
 Dim d As Dictionary
 Set d = ToD()
 Set medias = New Collection
 If param Is Nothing Then
  If Len(text) Then tmBotSend = tmBot(token, chat_id, "sendMessage", ToD("text", text))
  Exit Function
 Else
  If param.Count = 0 Then
   If Len(text) Then tmBotSend = tmBotForm(token, chat_id, "sendMessage", ToD("text", text))
   Exit Function
  End If
  filename = param.Keys
  pavd = param.Items  'pavd as "photo", "animation", "audio", "voice", "video", "video_note" and "document" as default
 End If
 sendChatAction = True 'only one sendChatAction
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
   If sendChatAction Then 'limits send per second
    Select Case send
    Case "photo", "voice", "video", "video_note", "document"
     tmBot token, chat_id, "sendChatAction", ToD("action", "upload_" & send)
     sendChatAction = False
    End Select
   End If
   If UBound(filename) Then 'group of media files
   Else
    tmBotSend = tmBotForm(token, chat_id, "send" & StrConv(Replace(send, "_n", "N"), vbProperCase), d)
    Exit Function
   End If
  End If
 Next
 d.Add "media", ConvertToJson(medias)
 tmBotSend = tmBotForm(token, chat_id, "sendMediaGroup", d)
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
 With New ADODB.Stream 'CreateObject("ADODB.Stream")
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
 If tmDeb Then Debug.Print "<" & filename & ">"
 With New ADODB.Stream 'CreateObject("ADODB.Stream")
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
 With New ADODB.Stream 'CreateObject("ADODB.Stream")
  .Mode = adModeReadWrite
  .Type = adTypeBinary
  .Open
  .write stringToBytes(send)
  For Each strA In fileC
   .write stringToBytes((strA(0)))
   .write fileToBytes((strA(1)))
  Next
  .write stringToBytes(bondS)
  .Position = 0
  If tmDeb = 2 Then .SaveToFile logFile, adSaveCreateOverWrite
  bodyToBytes = .read
  .Close
 End With
End Function

Function tmBotForm(token As String, chat_id As String, verb As String, param As Dictionary) As String
 'use ToD() to setup param
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
 With New MSXML2.XMLHTTP60 'CreateObject("MSXML2.XMLHTTP")
  .Open "POST", api & token & "/" & verb, False
  .setRequestHeader "Content-Type", "multipart/form-data; boundary=" & bond("", "")
  If tmDeb Then
   Debug.Print "POST " & api & "<token>/" & verb
   Debug.Print "Content-Type: multipart/form-data; boundary=" & bond("", "")
   Debug.Print
   T0 = Timer
  End If
  .send bodyToBytes(send, fileC, bond(suff:="--"))
  tmBotForm = .responseText
  If tmDeb Then
   Debug.Print
   Debug.Print ConvertToJson(ParseJson(.responseText), Whitespace:=1), Timer - T0
   WaitSec 'limits send per second
  End If
 End With
End Function
Function tmBot(token As String, chat_id As String, verb As String, param As Dictionary) As String
 'use ToD() to setup param
 send = api & token & "/" & verb & "?chat_id=" & chat_id
 For Each k In param.Keys
  send = send & "&" & k & "=" & WorksheetFunction.EncodeURL(param(k))
 Next
 With New MSXML2.XMLHTTP60 'CreateObject("MSXML2.XMLHTTP")
  .Open "POST", send, False
  If tmDeb Then
   Debug.Print "POST " & Replace(send, token, "<token>")
   T0 = Timer
  End If
  .send
  tmBot = .responseText
  If tmDeb Then
   Debug.Print ConvertToJson(ParseJson(.responseText), Whitespace:=1), Timer - T0
   WaitSec 'limits send per second
  End If
 End With
End Function
Function ToD(ParamArray param()) As Dictionary
 'for module JsonConverter from https://github.com/VBA-tools/VBA-JSON
 'add to project class Dictionary from https://github.com/timhall/VBA-Dictionary
 'or set a reference to Microsoft Scripting Runtime
 Set ToD = New Dictionary
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
 'message
 Set FirstMessage = ParseJson(tmBotSend(token, chat_id, "Ìû íà÷èíàåì ÊÂÍ"))
 tmBotSend token, chat_id, "Ïàïà", ToD()
 'https://core.telegram.org/bots/api#sendmessage
 tmBot token, chat_id, "sendMessage", ToD("text", "Ìàìà" & space(4096 - 8) & "Ïàïà")
 tmBotForm token, chat_id, "sendMessage", ToD("text", "Ìàìà" & space(4095 - 8) & "Ïàïà")
 
 'photo
 tmBotSend token, chat_id, "ôîòêà ïî ôàéë èä", ToD("AgACAgIAAxkDAANIY90VxfyqwbbEP7xy9MacV5VwcTAAAp_EMRtlgOlK8gV2JnFsXYcBAAMCAAN3AAMuBA", "photo")
 tmBotSend token, chat_id, "ôîòêà ïî ÓÐË", ToD("https://vremya-ne-zhdet.ru/wp-content/uploads/2020/04/picture174.png", "photo")
 'https://core.telegram.org/bots/api#sendphoto
 tmBot token, chat_id, "sendPhoto", ToD("caption", "ôîòêà ïî ôàéë èä", "photo", "AgACAgIAAxkDAANIY90VxfyqwbbEP7xy9MacV5VwcTAAAp_EMRtlgOlK8gV2JnFsXYcBAAMCAAN3AAMuBA")
 
 'attach photo
 tmBotSend token, chat_id, "âëîæåííàÿ ôîòêà", ToD("s:\01.jpg")
 'https://core.telegram.org/bots/api#sending-files
 tmBotForm token, chat_id, "sendPhoto", ToD("caption", "âëîæåííàÿ ôîòêà", "photo", "s:\01.jpg")
 
 'attach photo as document
 tmBotSend token, chat_id, "âëîæåííàÿ ôîòêà êàê ôàéë", ToD("s:\01.jpg", "document")
 'https://core.telegram.org/bots/api#senddocument
 tmBotForm token, chat_id, "sendDocument", ToD("caption", "âëîæåííàÿ ôîòêà êàê ôàéë", "document", "s:\01.jpg")
 
 'attach video as animation
 tmBotSend token, chat_id, "âëîæåííîå âèäåî êàê àíèìàöèÿ", ToD("s:\abaku.mp4", "animation")
 'https://core.telegram.org/bots/api#sendanimation
 tmBotForm token, chat_id, "sendAnimation", ToD("animation", "s:\abaku.mp4", "caption", "âëîæåííîå âèäåî êàê àíèìàöèÿ")
 
 'photos
 tmBotSend token, chat_id, "ôîòêè ïî ôàéë èä", ToD("AgACAgIAAxkDAANIY90VxfyqwbbEP7xy9MacV5VwcTAAAp_EMRtlgOlK8gV2JnFsXYcBAAMCAAN3AAMuBA", "photo", "AgACAgIAAxkDAANiY-HtiTrOf1yGJcU3_-9H2rwDLdEAAlXFMRuxTwlLqAge0lEC0wkBAAMCAAN5AAMuBA", "photo")
 'photos raw
 tmBot token, chat_id, "sendMediaGroup", ToD("media", ConvertToJson(Array(ToD("caption", "ôîòêè ïî ôàéë èä", "type", "photo", "media", "AgACAgIAAxkDAANIY90VxfyqwbbEP7xy9MacV5VwcTAAAp_EMRtlgOlK8gV2JnFsXYcBAAMCAAN3AAMuBA"), ToD("type", "photo", "media", "AgACAgIAAxkDAANiY-HtiTrOf1yGJcU3_-9H2rwDLdEAAlXFMRuxTwlLqAge0lEC0wkBAAMCAAN5AAMuBA"))))
 
 'attach photos
 tmBotSend token, chat_id, "âëîæåííûå ôîòêè", ToD("s:\01.jpg", "", "s:\02.jpg", "")
 'https://core.telegram.org/bots/api#sendmediagroup
 tmBotForm token, chat_id, "sendMediaGroup", ToD("media", "[{""caption"":""âëîæåííûå ôîòêè"",""type"":""photo"",""media"":""attach://01.jpg""},{""type"":""photo"",""media"":""attach://02.jpg""}]", "01.jpg", "s:\01.jpg", "02.jpg", "s:\02.jpg")
 
 'attach documents
 tmBotSend token, chat_id, "âëîæåííûå ôàéëû", ToD("s:\01.jpg", "document", "s:\02.jpg", "document")
 'attach documents raw
 tmBotForm token, chat_id, "sendMediaGroup", ToD("media", ConvertToJson(Array(ToD("caption", "âëîæåííûå ôàéëû", "type", "document", "media", "attach://p1"), ToD("type", "document", "media", "attach://p2"))), "p1", "s:\01.jpg", "p2", "s:\02.jpg")
 
 'attach photo video
 tmBotSend token, chat_id, "ôîòêà è âèäåî", ToD("s:\01.jpg", "", "s:\abaku.mp4", "")
 'attach photo video raw
 tmBotForm token, chat_id, "sendMediaGroup", ToD("media", ConvertToJson(Array(ToD("caption", "ôîòêà è âèäåî", "type", "photo", "media", "attach://p"), ToD("type", "video", "media", "attach://v"))), "p", "s:\01.jpg", "v", "s:\abaku.mp4")
 Set lastMessage = ParseJson(tmBotSend(token, chat_id, "Ðàñ÷¸ò îêîí÷åí"))
 Stop
 If Not FirstMessage("ok") Then Exit Sub
 If Not lastMessage("ok") Then Exit Sub
 If 1 Then
  First = FirstMessage("result")("message_id")
  Last = lastMessage("result")("message_id")
 Else
  First = 270
  Last = 212
 End If
 For i = First To Last 'https://core.telegram.org/bots/api#deletemessage
  Set deleteMessage = ParseJson(tmBot(token, chat_id, "deleteMessage", ToD("message_id", i)))
  Debug.Print i, deleteMessage("ok")
  If Not deleteMessage("ok") Then Debug.Print deleteMessage("description")
 Next
End Sub
