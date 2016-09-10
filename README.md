# excel-messenger
A silly local group messaging app where the server and all clients are excel workbooks.

Note: this is not a real app, just an excel demo of the client-server paradigm, see the [original blog post](http://tristancalderbank.com/2016/09/06/excel-messenger-a-terrible-experiment-in-vba/) that started this.

![folder](https://github.com/tristancalderbank/excel-messenger/blob/master/screenshots/cs-window.PNG?raw=true)

Features:
- Only works on a local network
- Clients all reference server sheet to get the master chat list
- Server references nicknames and messages from client worksheets, adding them to the master chat list and shifting everything up
- All sheets are autosaving constantly
- Messages can take over 40 seconds to be received

New user interface:
 
![folder](https://github.com/tristancalderbank/excel-messenger/blob/master/screenshots/cs-client-real.PNG?raw=true)
