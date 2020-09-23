----------------  
iQ Notepad® - Read Me File
---------------- 

A pretty significant update to the iQ Notepad clone project. iQ Notepad is a text based editor that maintains the basic functionality and integrity of Microsoft Notepad but adds enhancements like spell check, change case, importing/converting Rich Text and Microsoft Word files, inserting/appending files, a recent files menu, quick access to extended ASCII/ANSI symbols and characters, a toolbar and more. See the iQ NotePad Help Manual (PDF) accessible from the Help Menu of iQ Notepad. For those looking for a full featured word processor please see iQ WordPad also available on PSC. iQ WordPad is an easy to use Word Processor that includes features like Tables, Text Highlighting, Headers and Footers, Print Preview and many other features unavailable in Microsoft's Wordpad. Go to http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=69067&lngWId=1

This Code not tested in Win 98 or 95. Designed specifically for XP and tested in Vista and Windows 2000.

--------------------------------------------------------
          Update - Changes to iQ Notepad
--------------------------------------------------------

October 10th, 2007: Version 3.5

	" Added Case Change feature to Format Menu
	" Added file merge (insert a text file) option.
	" Added support for Spell Check
	" Added multiple Date/Time format options
	" Changed how window dimensions saved to ini file if program is ended while maximized or minimized.  Now starts maximized if last use was maximized.
	" Added Font and Background color. Aesthetic only for viewing in iQ Notepad. Files saved and printed in plain text with no formatting.
	" Added Print Preview to File Menu and Toolbar
	" GUI change to status bar and other minor GUI changes.
	" Added make default editor option to Help Menu.
	" Several minor bug fixes and code optimization

*************************************
July 1st, 2007: Version 2.12

	" Added tab default
	" Added error checking for Print/Page setup dialogs for systems
 with no printers.
	" Changed file load to check and display Unicode UTF-8 formatted files.
 Still limited to RTB showing ? for Unicode characters it doesn't read but
 now will display rest of text better. Changed loading a file to binary instead
 of LoadFile option of RTB.
	" Changed how window dimensions saved to ini file if program is ended while maximized or minimized.
	
	" Added and tested updated XP Theme resource file. No longer need external manifest file since this will compile into executable.

	" Have had requests for the compiled version of iQ Notepad including the menu ocx and help file. I have created an installation version of iQ Notepad. If you would like this version please email me and I will send it to you.

*************************************
June 29th, 2007:

	" Removed XP theme res file from project as it was creating problems for some when program was compiled. Replaced with iqnotepad.exe.manifest file. 

*************************************

-------------------------------
        Known Issues:
-------------------------------
1.  Unicode UTF-8 formatted files are detected and support is included for loading/viewing. Saving the file will change Unicode characters to a question mark (?). Included is a sample file named unicodeutf8.txt. The Unicode file will be displayed according to the locale set in the users version of windows. Actually, iQ NotePad displays this file more accurately than MS Notepad (compare the 2 side by side).

2.  The toolbars do not display properly on systems using Large Fonts (120 dpi). This will be addressed in a future version.

3.  There are some references to 3rd party spell-check and menu controls. These have been commented out.

-----------------------------------
Components Include:
-----------------------------------
Common Dialog: I know, I know. An unneeded dependency. I know there are lots of replacement modules out there. The fact is that I slapped this on when starting the project and was too lazy to replace it by the time I thought about it. Shouldn't be too much work to replace that if you want.

Candy Button Control: I added this control from Mario Villanueva's award winning control (it's available at http://www .planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=64969&lngWId=1). Used throughout the project and especially prominent part of the Toolbar. Thank you Mario!

Flex-Grid: This component is used for the Symbols dialog box feature of iQ Notepad.

Toolbar: The VB6 Toolbar control has such a 1998 look to it. Don't like it much. So, created a picture box and used the Candy Buttons for the toolbar.

Status Bar: The status bar in Microsoft's Notepad displays only if word wrap is off. The only status it provides is the line and column number of the cursor location. For the iQ status bar we have a picture box, a label and another picture box for the grabber image. I programmed the status bar to be independent of the word wrap status.

XP Resource File: Included so that compiled program will have XP Theme on applicable operating systems. 

Rich Textbox Control: Required because of the file size limitations of the Textbox control.  Using this control in a text only project presents many challenges. One is in printing from this control. Another is formatting. The Rich Text control's whole purpose is to display stuff formatted with fonts, colors and graphics. Suppressing this stuff to display text with no formatting created some problems. This is especially true with the edit functions, in particular pasting from the clipboard. The Shortcut keys for this control (Ctrl-V for paste) are built into the control and are NOT trappable in the Key Preview event of your main form.  And what happens if you assign the same Shortcut Keys in your Edit Menu using the menu editor? The result is that it will double paste everything. There are a couple of ways to approach this... one would be sub classing, hooking/unhooking, etc. Didn't want to go there. You'll see the approach used when viewing the code. May not be pretty but it does the job.

INI File: Microsoft Notepad auto saves your font attributes, printing page setup options and the size of your window. iQ Notepad does the same thing using an ini file.  Could have done this to the registry but I don't like writing to the registry file unless necessary. Recent Files also written to the ini file. All that code is in the modIni.bas module.

Command Line and Drag/Drop: iQ Notepad will auto open any file specified on a command line. If you associate iQ Notepad as your default text editor that means when you double click on a text file in Windows Explorer iQ Notepad will open displaying that file. You also may drag/drop onto iQ Notepad to view files.

Spell Check: This feature requires the user have some version of Microsoft Office or Word on their system. What happens here is that it actually launches an instance of Word using VB's CreateObject. Word is invisible and actually placed off screen. Spell check will go through the document, or selected text, and the spelling dialog box will be displayed. If no errors are detected, or when checking is finished, a message box will appear that the spell check is complete. Grammar checking is available in this object. If you want to include grammar check just remove comment on the line that says ".Checkgrammar".

--------------------------------
     Components Not Used
--------------------------------
In the project I created for my client I used a 3rd party menu control, XPNetMenu. You can see how it looks in the Screen shot. Since this is a commercial control it was removed from this project. If you're looking for a great menu control at a reasonable price check them out at http://www.xpstyle-menu.com. I think they even sell the VB6 source code for this control.

In my compiled version I also have a 3rd party spell check control.  That code has been commented out and code to use MS Word spell check is active.  This does require you have some version of Word installed on your system.

For those looking for an enhanced version of Wordpad please see iQ WordPad also available on PSC. iQ WordPad is an easy to use Word Processor that includes features like Tables, Text Highlighting, Headers and Footers, Print Preview and many other features unavailable in Microsoft's Wordpad.

That's basically it. If you would like a compiled version of iQ NotePad with all 3rd party controls and features active let me know and I'll email to you.  If you have any questions or wish to communicate with me directly you may email me at tmoran4511@hotmail.com.

Have fun!


