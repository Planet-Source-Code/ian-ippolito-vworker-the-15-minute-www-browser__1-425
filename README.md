<div align="center">

## The 15 minute WWW browser


</div>

### Description

It took Netscape months to create their first browser, and Microsoft wasnt able to follow up that trick until years later. But you can create a world wide web browser in less than 15 minutes, even if you are a Visual Basic novice! The following tutorial will show you how to create an Internet browser using the Microsoft Internet Control (part of Internet Explorer). My only caveat is that if you end up putting the big-boys out of the Internet business with your creation, that you give me a litle credit in your About box! ;)
 
### More Info
 
If you are running Visual Basic 5.0 or 6.0 the Microsoft Internet Controls are included with VB, so you don't need to do anything special to get them. If you are running an older version, you may still be able to get the controls for free from Microsoft's site at www.microsoft.com.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Ian Ippolito \(vWorker\)](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/ian-ippolito-vworker.md)
**Level**          |Beginner
**User Rating**    |3.9 (31 globes from 8 users)
**Compatibility**  |VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Internet/ HTML](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/internet-html__1-34.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/ian-ippolito-vworker-the-15-minute-www-browser__1-425/archive/master.zip)





### Source Code

```
1)Create a new Visual Basic project.
2)Add the Microsoft Internet Controls to your project. In
VB6 you add new custom controls to a project by going to
the Project menu and choosing the Components sub-menu, and
choosing the control you want to add. In other versions of
VB, consult your help on adding custom controls. The name
of the custom control is: Microsoft Internet Control. This
will add two icons to your toolbox. Place the one that
looks like a globe (the Web Browser control) on your form
by double-clicking it. This control will display the web
page, so make sure you size it so that it looks presentable.
3)Next, place a text box on the upper portion of the form--
above the WebBrowser Control. This will be your browser's
address bar. To complete the address bar, place a button
next to it. Change the Caption property of the button to:&Go
4)Now add the following code to your form:
Private Sub Command1_Click()
 WebBrowser1.Navigate Text1
End Sub
That is it! Run your project and type www.microsoft.com
into the text box and press the GO button. (Dont forget to
start your Internet connection if its not already up). The
page will load and display just like a browser!
Now that you have an idea of how simple the control is to
use, you can take a little more time to create some more
sophisticated functionality for your browser:
1)Since the world wide wait can be taxing on your browser
users, you can create a status bar at the bottom of your
form that lets them know how much of their page has loaded.
You can use the following web browser events (see the
Microsoft Internet Controls help file, if you need examples)
WebBrowser1_DownloadBegin
WebBrowser1_DownloadComplete
WebBrowser1_ProgressChange
2)Create a menu system on your form--just like IE and
Netscape. See the VB help if youve never done this before.
You'll want to at least create &File and &Exit.
3)Create a combobox instead of a text box that remembers
old URLs.
4)Let your imagination run wild!
5) For more features, check out the other browser
submissions to this site. An outstanding example is:
http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=2628
6)Some other features created by other users:
Passive Matrix(mailto:Passive_matrix@hotmail.com)
Here are some helpful button commands...
back button
WebBrowser1.GoBack
Forward button
WebBrowser1.GoForward
refresh
WebBrowser1.Refresh
stop
 WebBrowser1.Stop
home
WebBrowser1.Navigate ("www.cow.com")
By William:mailto:wfloor@rendo.dekooi.nl
An answer to the questions about the favorites and the bookmarks:
1) Make a
commandbutton cmdAdd
2) Make a commandbutton cmdFav
3) Make a listbox
lstFavs
The code for cmdAdd:
Private Sub cmdAdd_Click()
 FN =
FreeFile
 Open "favs.txt" For Append As FN
 Print #FN, txtUrl.Text &
Chr(13)
 Close #FN
End Sub
The code for cmdFav:
Private Sub
cmdFav_Click()
 On Error Resume Next
 FN = FreeFile
 Open
"favs.txt" For Input As FN
 lstFavs.Visible = True
 Do Until
EOF(FN)
  Line Input #FN, NextLine$
  lstFavs.AddItem NextLine$
 Loop
 Close #FN
End Sub
The code for lstFavs:
Private Sub
lstFavs_Click()
 txtUrl.Text = lstFavs.List(lstFavs.ListIndex)
txtUrl_KeyPress 13
 lstFavs.Visible = False
 Close #FN
End Sub
By:CheaTzZ mailto:cheatzz@xcheater.com
To print:
Private Sub printmenu_Click()
 Dim eQuery As OLECMDF
 On
Error Resume Next
 eQuery = WebBrowser1.QueryStatusWB(OLECMDID_PRINT)
If Err.Number = 0 Then
  If eQuery And OLECMDF_ENABLED Then
WebBrowser1.ExecWB OLECMDID_PRINT, OLECMDEXECOPT_PROMPTUSER, "", ""
Else
   MsgBox "The Print command is currently disabled."
End If
 Else
  MsgBox "Print command Error: " & Err.Description
End If
End Sub
======================
To open up new window:
Private
Sub WebBrowser1_NewWindow2(ppDisp As Object, Cancel As Boolean)
 On Error
Resume Next
 Dim frmWB As Form1
 Set frmWB = New Form1
 Set
ppDisp = frmWB.WebBrowser1.Object
 frmWB.Visible = True
 Set frmWB =
Nothing
If you want to cancel the new window, Cancel = True.
For a proper progressbar:
Private Sub WebBrowser1_ProgressChange(ByVal
Progress As Long, ByVal ProgressMax As Long)
 On Error Resume Next
ProgressBar1.Max = ProgressMax
 ProgressBar1.Value = Progress
End
Sub
To show the percentage:
Progress * 100 / ProgressMax
by: Bones mailto:kacantu@webaccess.net
You can easily view the source of the webpage you're
viewing by using 2
controls: A RichTextBox control,
and the microsoft internet transfer
control.
If your internet transfer control is Inet1, and your
Textbox is
RichTextBox1, then use the following code
download and view a page's
source:
RichTextBox1.Text = Inet1.OpenURL(" address ")
The address must be
the valid URL of an .htm or .html
file.
```

