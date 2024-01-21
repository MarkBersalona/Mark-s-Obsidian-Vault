# Install VB6 on Windows 7
Created: 2024.01.21 1701


Keywords: VB6 Visual Basic 6 Visual Studio Install How To Tutorial Legacy Development
Tags: Visual Basic 6 Tutorial  Windows 7 Vista
Entered On: 2009-06-23
Views: 29407 

After surfing around the net, I've found very little information regarding installation of VB6 on Windows 7. Most of the information out there is for Vista, and most of it is queries for assistance.

You may be wondering why someone would want to utilize VB6 on a shiny new operating system like Windows 7. Or even Vista for that matter.

There are about a bazillion legacy applications out there that have to be supported, and people like me who speak VB6 need to have the tools installed on our workstations in order to implement and test updates and such for these legacy applications. 

It also helps out when I need to squirt out a quick tool for use in my daily work.

So without further delay, here is the process that I have used on my Windows 7 machines to install Visual Basic 6.

Bonus tip from BadBrad: I forgot to post about this part, but BadBrad reminded me in the comments.

Before proceeding with the installation process below, create a zero-byte file in C:\Windows called MSJAVA.DLL. The setup process will look for this file, and if it doesn't find it, will force an installation of old, old Java, and require a reboot. By creating the zero-byte file, the installation of moldy Java is bypassed, and no reboot will be required. Thanks for the reminder, BadBrad! 

  1. Turn off UAC.

  2. Insert Visual Studio 6 CD.

  3. Exit from the Autorun setup.

  4. Browse to the root folder of the VS6 CD.

  5. Right-click SETUP.EXE, select Run As Administrator.

  6. On this and other Program Compatibility Assistant warnings, click Run Program.

  7. Click Next.

  8. Click "I accept agreement", then Next.

  9. Enter name and company information, click Next.

 10. Select Custom Setup, click Next.

 11. Click Continue, then Ok.

 12. Setup will "think to itself" for about 2 minutes. Processing can be verified by starting Task Manager, and checking the CPU usage of ACMSETUP.EXE.

 13. On the options list, select the following:
	   · Microsoft Visual Basic 6.0
	   · ActiveX
	   · Data Access
	   · Graphics

All other options should be unchecked. Click Continue, setup will continue.

 14. Finally, a successful completion dialog will appear, at which click Ok. At this point, Visual Basic 6 is installed.

 15. If you do not have the MSDN CD, clear the checkbox on the next dialog, and click next. You'll be warned of the lack of MSDN, but just click Yes to accept.

 16. Click Next to skip the installation of Installshield. This is a really old version you don't want anyway.

 17. Click Next again to skip the installation of BackOffice, VSS, and SNA Server. Not needed!

 18. On the next dialog, clear the checkbox for "Register Now", and click Finish.


The wizard will exit, and you're done. You can find VB6 under Start, All Programs, Microsoft Visual Studio 6.

I've not tested this on Vista, but it should work for you there as well. Give it a try, and let me know how it works!

Enjoy!

--------------------------------------------------------------------------------------

UPDATE

You might notice after successfully installing VB6 on Windows 7 that working in the IDE is a bit, well, sluggish. For example, resizing objects on a form is a real pain.

After installing VB6, you'll want to change the compatibility settings for the IDE executable.

  1. Using Windows Explorer, browse the location where you installed VB6. By default, the path is C:\Program Files\Microsoft Visual Studio\VB98\

  2. Right click the VB6.exe program file, and select properties from the context menu.

  3. Click on the Compatibility tab.

  4. Place a check in each of these checkboxes:
	   · Run this program in compatibility mode for Windows XP (Service Pack 3)/li>
	   · Disable Visual Themes/li>
	   · Disable Desktop Composition

	   · Disable display scaling on high DPI settings



After changing these settings, fire up the IDE, and things should be back to normal, and the IDE is no longer sluggish.

Content Thieves:
The below sites have stolen the content of this page and posted as their own work. Congratulations, losers. You've discovered the power of copy/paste.

OpenSourceCoimbatore
hackingaday

=========================================================================

Comments On This Post

By: earlymeadow
Date: 2009-11-15

Super information - got VB6 up and running, your explanation saved me a lot of searching.

I'm still having a few problems with the INet transfer control - using it to transfer files via FTP, it seems to hang for a minut or so every now and then. 

FWIW - using manifest files (as in XP) works in Windows 7 - giving the app the windows 7 look.

Thanks again.

Peter

By: FortyPoundHead
Date: 2009-11-16

Thanks, Peter! Great tip about the manifest files.

By: badbrad
Date: 2009-11-20

Missed one thing.... Create a zero byte file in C:Windows called "MSJAVA.DLL" to get around the Java install.

Also if you install the "VB6IDEMouseWheelAddin.dll" for scrolling, make sure that you run "cmd.exe" with elevated admin privleges when registering the DLL manually. Also ran it with the DLL in "c:".

Sorry... no links for this right now.

By: mccindy
Date: 2010-01-09

Thank you so much - I will be attemping this isnatll very soon. 

By: nfordatfph
Date: 2010-02-09

Thanks for the tip on fixing slugishness. It was driving me crazy.

I have a problem with large forms in the IDE in Win7. It gets to something like 1000x800 and won't let me drag it any bigger. At runtime, I can set the form to a larger size, but not being able to see the full form in the IDE makes it pretty hard to work on.

By: FortyPoundHead
Date: 2010-02-09

A little odd, but I'll see if I can reproduce it here, and maybe find a fix for you.

By: FortyPoundHead
Date: 2010-02-10

I haven't been able to reproduce the issue here. I'm able to drag the form size to a much larger size than the screen, which is currently 1280x1024. 

I've not seen this before - anyone got any ideas?

By: pamith
Date: 2010-02-17

What about third party OCX applications? I have a program that uses several of these. I want to move my VB6 work to Windows 7 but would need to know if the third party OCXs will work as well. Suggestions?

By: FortyPoundHead
Date: 2010-02-17

As long as the OCX's get registered properly, there should be no issue. That is, unless the OCX is doing something *really* strange, you shouldn't have a problem.

Let me know what OCX's you are concerned about, and I'll try testing them out on my installation, and see if they work.

By: caesarv
Date: 2010-02-17

pardon my ignorance...but...
I just got Win7 Pro 64 installed on my new computer last night. Most of my work is in VB6. Will the instructions shown work with 64 bit Windows? I have heard that VB6 IDE is not supported with 64bit. I have the original VB6 CD but, of course, it does not have SP6. Is this a concern or can I simply add SP6 after installation. FYI, I do have access to an msdn subscription if I need to download something different.

By: FortyPoundHead
Date: 2010-02-17

Yes, it will work fine on 64 bit Windows 7. I am currently using 64 bit on my main workstation, and have had no issues. 

Applying SP6 after installation of VB6 has posed no challenges, either.

By: PhilTilson
Date: 2010-03-04

Excellent - thank you! It all worked just as you said - and .exe files run quite happily as well - though with a considerable 'starting delay'!

By: mgpprasan
Date: 2010-03-28

what a useful information. thanks a lot. i was suffering from how to install vb 6.0 in windows 7. now it is working well. 

By: FortyPoundHead
Date: 2010-03-29

Glad to be assistance!

By: George
Date: 2010-04-28

I have an issue getting the XP manifest file to work in WIndows 7 when I run VB in IDE mode. It does not look like XP. I think it might have to do with permissions with either VB6, the manifest file. Any ideas?

By: vic
Date: 2010-06-04

I have a fairly large VB6 App that I've been maintaining in Windows 7 using Virtual PC. I decided to try to get it to work in Windows 7 with Virtual PC and there are a number of problems. I've gotten around each one but this one now has me stumped.

The app uses an Access Database and has lots of forms using ADO data controls. What happens on some of the forms (not all of them) is when I tryh to do a requery on an a recordset object i.e. Me.Adodc1.Recordset.Requery, it just hangs (no message, nothing). Then when I try to close the app I get the message "Visual Basic has stopped working" and it's gone.

Any ideas?

Vic

By: jnc39
Date: 2010-06-21

As a brand new Windows 7 user, I have a simple question: What is the "UAC" you refer to in step1? Could the fact that I didn't shut it off before starting your process have caused the computer to lock up and require a hard shut down at step 10?
Thanks for your help

By: FortyPoundHead
Date: 2010-06-21

UAC is user account control, which helps protect your computer from unauthorized changes. You can turn it off in the control panel.

By: FortyPoundHead
Date: 2010-06-22

Guess it would help I posted the information about disabling UAC:

http://www.fortypoundhead.com/showcontent.asp?artid=19521


By: anotherhacker
Date: 2010-06-23

It tried installing a program that I created in VB6 on a W7 machine and I get an "Invalid Line In Setup Information File" errror on the install.....?

By: FortyPoundHead
Date: 2010-06-23

I am currently attending HP techforum in Las Vegas, and as such am currently experiencing limited connectivity. But I'll look into this for you when I get back to my room.

By: FortyPoundHead
Date: 2010-06-25

AnotherHacker: have you seen:

http://support.microsoft.com/kb/216156

It directly addresses the issue you are experiencing.

By: Jovick
Date: 2010-06-28

Great article! I have installed VB6 on Win 7 64-bit successfully. Some things I ran into that I didn't see mentioned is that when registering OCX and DLL files the file MUST be placed in the C:WindowsSysWOW64 directory. I know, I know... it seems the reverse of what you should do. You MUST also use the regsvr32.exe file in the SysWOW64 folder (not the one in the System32 folder) to register the files. One other tip is to ALWAYS install everythig as "administrator". I found that I also had to modify the properties of the VB6.EXE file (in the compatibility tab) to "run as administrator".

By: FortyPoundHead
Date: 2010-06-29

Thanks for the great tip, Jovick! 

By: raycomp
Date: 2010-07-10

Thanks for ht useful information. One problem I sometimes need to register 3rd party ocx files but in most cases Windows 7 32bit does not allow copying them to Windowssystem32 for registering but place the in the incompatible folder. Any ideas please

By: FortyPoundHead
Date: 2010-07-10

Need some more info from you, Raycomp. Are you simply copying the files with Windows explorer? Or is the file copy part of a setup process?

If it is part of your application set up, you might try running the setup program as administrator - right click the setup file, and select Run As Administrator.

By: raycomp
Date: 2010-07-10

Thanks for the quick reply. I have a few minutes ago found the answer and was about to post the fact here when I saw your post.

I am developing an app to keep track of all my movie files ant to also play them with either WMP or Gom from inside the App. All ok. 

By: mwrob
Date: 2010-07-20

I followed the instructions and was able to get VB6 installed just fine on a Windows 7 64 bit system. However, I am unable to get service pack 6 installed.

I have set the properties of acmesetup and setupsp6 to run as administrator and Windows SP 3 mode. Right after agreeing to the license terms, I get a dialog box that says "Visual Studio 6.0 Service Pack 6 Setup was not completed properly."

By: FortyPoundHead
Date: 2010-07-20

Have you turned off UAC during the installation? 

I just tried in a test environment, utilizing Windows 7 64 bit, installed VB6 using the above procedure, then downloaded the VB6 SP6 from the Microsoft site (direct download from MS).

Double clicking the executable allowed me to select an temp folder to extract the SP6 files to. After extraction, I browsed to the directory where the files extracted to, and ran the SetupSP6.exe application. 

The installation went normally, with no errors, and the final dialog stating setup was successful.

However, after resetting the environment, I was able to reproduce the error you mention by leaving UAC turned on.

By: mwrob
Date: 2010-07-21

Yes, I did exactly as you described. I tried again this morning, double checking all the settings and still got the same result. I'm open to suggestions. Thanks.

By: FortyPoundHead
Date: 2010-07-21

I'm currently out of town, but I'll check into it more this evening.

By: mwrob
Date: 2010-07-23

Ok, suppose I can't get SP 6 installed. What do I need to beware of if I try working on a project last worked on in SP6 and I work on it in VB6 with no service packs installed?

By: latino
Date: 2010-08-30

Hi,

I have VB6 working perfectly. Sourcesafe works too, but I can't work with sourcesafe from within the IDE. The SS add-in is loaded, there are no errors shown, but I don't have any extra contextmenu-items. Anyone here had the same problem?

By: philyuko
Date: 2010-09-01

Great info for Windows 7 … works like a dream.
Thank you!!!
For Vista - there is no service pack 3 so I clicked 2.
For the most part it works OK … but two hiccups

1/ From time to time you’ll click SAVE AS and no box will come up
2/ When a save box does come up it doesn’t have the name of the file in it … in other words … I don’t think this works for vista … any ideas?

By: FortyPoundHead
Date: 2010-09-01

Are you using the common dialog control?

By: philyuko
Date: 2010-09-01

No ... but I'd like to. I had to take out all the CD controls because it meant my users without XP had to do too much installing of too many things in sys32. I'm sure you know what I mean.

By: FortyPoundHead
Date: 2010-09-02

I'll hammer on this today, and see if I can duplicate the behavior. Are you getting any weird event log entries that might be related? Along the lines of Path not found or access denied, or some such?

By: philyuko
Date: 2010-09-03

Nope ;-)
Just won't won't give me the current name of the form or project when using save-as.

Could it be related to not have "Service pack 3"?

By: FortyPoundHead
Date: 2010-09-04

MMM... prolly not. Have you installed SP6 for VB6 yet?

By: philyuko
Date: 2010-09-04

I have not ... I should?
Any idea why that would a thang for Vista and not for Windows 7?


By: FortyPoundHead
Date: 2010-09-05

To put it simply, Vista is a bit more picky than Windows 7. That's why you didn't see businesses rolling Vista out to their employees, and IT folks every made the sign of the cross whenever they had to support it.

Give the SP6 a try and you might see your problems clear up.

By: psyman
Date: 2010-10-05

I get to the stage during install where you select custom, then I get an error saying I dont have enought memory to install. Clearly I do as I have 2GB ram & 200+GB free hard drive space. Any ideas?

By: FortyPoundHead
Date: 2010-10-05

First, make sure your system date and time are correct. If that doesn't fix the problem, or if your date and time are good to go, you can try pre-installing the MS VB6 Common Controls *before* installing VB6. 

You can grab the installation package from here.

By: psyman
Date: 2010-10-05

Thank you - that did it ! I now have VB6 again & it works a treat!

By: FortyPoundHead
Date: 2010-10-05

Glad I could help - don't hesitate to ask back here if you have any other questions!

 

You must be logged in to post a comment.



## References/See also:


## Tags
