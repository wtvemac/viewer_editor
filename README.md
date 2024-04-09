# WebTV Viewer Editor

This is an old tool I created that edits the WebTV Viewer to allow you to connect to a server like the box. It applies mods to the exe to bypass stops from Microsoft preventing you to connect to server. It also modifies the request header sent to the server so the viewer looks like a box rather than a viewer.

Around 2000, I switched to a Windows PC but still hacked around with the WebTV. The WebTV viewer was an important tool to play around on the server. This tool helped me explore and reverse engineer WNI's servers. For better or worse, it was also used by MattMan and others to hack into accounts since you could easily modify your SSID and use my [WorkAround.pl](https://github.com/wtvemac/Emu15/blob/main/WorkAround.pl "WorkAround.pl") server to push a ticket to the viewer and then redirect into someone's account.

- `Bin/WebTVIPE--4.0.pl` is the original Perl script I made before the exe.
- `Bin/SuperViewerIPE_4.0.exe` is the executable I made after the Perl script that allows you to make more mods in an easier way than the Perl command line script. 
- `Bin/SuperViewerIPE_4.1.exe` is just the 4.0 exe with mods from MattMan. He removed the ability to edit specific headers individually and just added a textbox.
- `Bin\WebTV-Viewer\WebTVIntel--1.0-consoleout.exe` as a bonus, this is the 1.0 viewer with the subsystem changed to a command line app so you can see STDERR and STDOUT messages in the console.


There's a basic template system I built to describe the type of edits this tool should do. You can find some of the templates inside the `Config/` directory.

There's also other helper utilities built into this tool. Like the ability to extract information out of the SSID and generate a username and password used to hack into WNI's 1800 servers.

This is meant for archival purposes. This hasn't been worked on in years. It's here for you to see my excellent programming and spelling skills.

I'd suggest you take a look at https://turdinc.kicks-ass.net/Msntv/viewer/Hackers-Edition-WebTV-Viewer-Auto-Gen-SSID.html if you want something more modern.
