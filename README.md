# tmBot
SendX where X: Message, Photo, ... MediaGroup by Telegram Bot API from VBA
## Credits
- Telegram - for [@BotFather](https://t.me/BotFather)
- VasiliY_Seryugin - for [Sub telegram_send_picture()](https://www.planetaexcel.ru/forum/index.php?PAGE_NAME=message&FID=1&TID=93149&TITLE_SEO=93149-kak-sdelat-otpravku-v-telegram-iz-makrosa-vba-excel&MID=1193376#message1193376) as father of tmBot
- Tim Hall - for [VBA-JSON](https://github.com/VBA-tools/VBA-JSON)
- Tim Hall - for [VBA-Dictionary](https://github.com/timhall/VBA-Dictionary) 
## Usage
- Add to your project module [tmBot.bas](tmBot.bas)
- Add to your project module [JsonConverter.bas](https://github.com/abakum/VBA-JSON/blob/master/JsonConverter.bas)
- Set Public Const token=<bot_id>:<bot_password>
- Set Public Const chat_id=<chat_id>
- Run test()
