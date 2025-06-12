1. VBA MAcro

Open Outlook.

Press ALT + F11 to open the VBA Editor.

Insert a new module: Right-click Project1 → Insert → Module.

Paste the code of "outlook_response.vba"

In Outlook, go to File → Options → Customize Ribbon.

Under the right-side list, pick a tab (e.g., Home) and click "New Group."

With the new group selected, click "Choose commands from: Macros."

Find Project1.GenerateReplyWithPython and click "Add."

Rename and assign an icon if you want.


2. Security/Trust Center Settings

Go to File > Options > Trust Center > Trust Center Settings > Macro Settings

Ensure "Notifications for all macros" is selected. 

If you click the Ribbon button configured above, a notification will open, asking you to activate Macros. After this, the AI-Assistant should work.