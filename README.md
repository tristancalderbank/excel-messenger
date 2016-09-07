# excel-messenger
A silly local group messaging app where the server and all clients are excel workbooks.

![folder](https://github.com/tristancalderbank/excel-messenger/blob/master/screenshots/cs-window.PNG?raw=true)

Features:
- Only works on a local network
- Clients all reference server sheet to get the master chat list
- Server references nicknames and messages from client worksheets, adding them to the master chat list and shifting everything up
- All sheets are autosaving constantly
- Messages can take over 40 seconds to be received! Wow!
- If you have another unrelated workbook open, sometimes the code starts writing messages into your workbook instead of the client one (this is handy if you don't want to switch back and forth)
- Very strong security, anyone can edit anyone else's sheet or even the server sheet, that way everyone can be on the lookout for hackers

New user interface:
 
![folder](https://github.com/tristancalderbank/excel-messenger/blob/master/screenshots/cs-client-real.PNG?raw=true)
