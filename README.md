<div align="center">

## Using the WebBrowser Control as an HTML Editor


</div>

### Description

Whether it be creating your own mail client or HTML editor, the WebBrowser control contains a lot more functionality than first appears. With a few simple tricks the WebBrowser control can be turned into a fully featured editor.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Roderick Thompson, CebraSoft](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/roderick-thompson-cebrasoft.md)
**Level**          |Intermediate
**User Rating**    |4.9 (74 globes from 15 users)
**Compatibility**  |VB 6\.0
**Category**       |[Internet/ HTML](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/internet-html__1-34.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/roderick-thompson-cebrasoft-using-the-webbrowser-control-as-an-html-editor__1-42025/archive/master.zip)





### Source Code

<body>
<p align="center"><font face="Arial" size="5">Using the WebBrowser Control as an
HTML Editor</font></p>
<p align="center"> </p>
<p><font face="Arial">This article is a simple introduction to how you can use
the WebBrowser control as an HTML editor. There are many features available
however there is no sample documentation and what little there is is in C++. </font></p>
<p><font face="Arial">The webbrowser control is added using Project, Components,
Microsoft Internet Controls and appears on the toolbox as a globe icon.<br>
<br>
When working with this control it is important to ensure that you initialise it
otherwise you will get Run Time Error 438.<br>
To initialise the control as an editable control enter something like the following:<br>
<br>
<i><b><font size="2">Private Sub Form_Load<br>
    WebBrowser1.Navigate2 "about:blank"<br>
    WebBrowser1.Document.DesignMode = "On"<br>
End Sub</font></b><br>
</i><br>
Now that you can type into the control, you will more than likely want to enter
rich text ie bold, italic, underline etc.<br>
<br>
By using the ExecCommand within the Document model, you can pass through
commands directly into the document. <br>
For example:<br>
<br>
<i><b><font size="2">Private Sub cmdBold_Click<br>
    WebBrowser1.Document.Execcommand "Bold"<br>
End Sub</font></b></i><br>
 </font></p>
<p><font face="Arial" size="4"><u><b>Basic Commands</b></u></font></p>
<p><font face="Arial">In the same way you can see in the example above, the
following commands can be sent straight through to the WebBrowser. <u><b><br>
<br>
</b></u><i><b><font size="2">WebBrowser1.Document.Execcommand "JustifyLeft"<br>
WebBrowser1.Document.Execcommand "JustifyCenter" <br>
WebBrowser1.Document.Execcommand "JustifyRight"<br>
WebBrowser1.Document.Execcommand "Bold"<br>
WebBrowser1.Document.Execcommand "Italic"<br>
WebBrowser1.Document.Execcommand "Underline"<br>
WebBrowser1.Document.Execcommand "Copy"<br>
WebBrowser1.Document.Execcommand "Cut"<br>
WebBrowser1.Document.Execcommand "Paste" <br>
WebBrowser1.Document.Execcommand "InsertHorizontalRule"<br>
WebBrowser1</font></b></i></font><b><i><font face="Arial" size="2">.Document.Execcommand
"Indent"<br>
</font></i></b><font face="Arial"><i><b><font size="2">WebBrowser1</font></b></i></font><b><i><font face="Arial" size="2">.Document.Execcommand
"Outdent"</font></i></b><font face="Arial"><br>
 </font></p>
<p><u><b><font face="Arial" size="4">Commands with Parameters</font></b></u></p>
<p><font face="Arial"><u><i><b>Font Size<br>
</b></i></u><br>
When looking to modify the font properties, parameters will need to be
specified.<br>
<br>
In the case of setting Font Size you need to use HTML sizes rather than point
sizes as follows:</font></p>
<div align="center">
 <center>
 <table border="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" id="AutoNumber1">
 <tr>
  <td width="50%" align="center"><font face="Arial"><b>HTML Size</b></font></td>
  <td width="50%" align="center"><font face="Arial"><b>Traditional Point
  Size</b></font></td>
 </tr>
 <tr>
  <td width="50%" align="center"><font size="4" face="Arial">1</font></td>
  <td width="50%" align="center"><font face="Arial" size="4">8pt</font></td>
 </tr>
 <tr>
  <td width="50%" align="center"><font size="4" face="Arial">2</font></td>
  <td width="50%" align="center"><font face="Arial" size="4">10pt</font></td>
 </tr>
 <tr>
  <td width="50%" align="center"><font size="4" face="Arial">3</font></td>
  <td width="50%" align="center"><font face="Arial" size="4">12pt</font></td>
 </tr>
 <tr>
  <td width="50%" align="center"><font size="4" face="Arial">4</font></td>
  <td width="50%" align="center"><font face="Arial" size="4">14pt</font></td>
 </tr>
 <tr>
  <td width="50%" align="center"><font size="4" face="Arial">5</font></td>
  <td width="50%" align="center"><font face="Arial" size="4">18pt</font></td>
 </tr>
 <tr>
  <td width="50%" align="center"><font size="4" face="Arial">6</font></td>
  <td width="50%" align="center"><font face="Arial" size="4">24pt</font></td>
 </tr>
 <tr>
  <td width="50%" align="center"><font size="4" face="Arial">7</font></td>
  <td width="50%" align="center"><font face="Arial" size="4">36pt</font></td>
 </tr>
 </table>
 </center>
</div>
<p><font face="Arial">In converting the font sizes to a value between 1 and 7,
the command is as follows (note the size is passed across as a string, not a
number):</font></p>
<p><font face="Arial"><b><font size="2"><i>WebBrowser1.Document.Execcommand "FontSize",
"", "5"<br>
<br>
</i></font><i><br>
<u>Font Color</u></i></b><br>
<br>
In the same way we had to use the HTML value for size rather than the windows
value, so too do we have to do the same for colors. We cannot use the Windows
long values and need to convert to the web format of RGB.<br>
<br>
I have modified a function from PlanetSourceCode that does the job although I
suspect there is a far easier way of using it!<br>
</font><font face="Arial" size="2"><b><i><br>
Public Function GetHexColor(theColor As Long) as String<br>
    Dim Red As Integer, Green As Integer, Blue As Integer<br>
<br>
    Red = theColor Mod &H100: theColor = theColor \ &H100<br>
    Green = theColor Mod &H100: theColor = theColor \ &H100<br>
    Blue = theColor Mod &H100<br>
<br>
    If Len(Hex(Red)) = 1 Then GetHEXColor = GetHEXColor & "0"<br>
    GetHEXColor = GetHEXColor & Hex(Red)<br>
    If Len(Hex(Green)) = 1 Then GetHEXColor = GetHEXColor & "0"<br>
    GetHEXColor = GetHEXColor & Hex(Green)<br>
    If Len(Hex(Blue)) = 1 Then GetHEXColor = GetHEXColor & "0"<br>
    GetHEXColor = GetHEXColor & Hex(Blue)<br>
    GetHEXColor = "#" & GetHEXColor<br>
End Function<br>
<br>
</i></b></font><font face="Arial">This function allows you to take colors from
the Command Dialog control and convert to web based colors. A simple
implementation would be as follows (taking note the format of the data passed
through to the command ie #RRGGBB):</font></p>
<p><i><b><font face="Arial" size="2">On Error Resume Next<br>
CommonDialog1.CancelError = True<br>
CommonDialog1.ShowColor<br>
If err.Number = 0 Then<br>
    HexColor = GetHexColor(CommonDialog1.Color)<br>
    WebBrowser1.Document.Execcommand "ForeColor", "", HexColor<br>
End If<br>
On Error Goto 0</font></b></i></p>
<p><font face="Arial"><b><i><u><br>
Font </u></i></b><u><b><i>Name</i></b></u><br>
<br>
As with the previous examples, the font name is exactly the same as the other
methods where the name of the font is passed as a string:</font></p>
<p><i><b><font face="Arial" size="2">WebBrowser1.Document.Execcommand "FontName",
"", "Arial"</font></b></i></p>
<p> </p>
<p><u><font face="Arial" size="4"><b>Reading and Writing to/from the Control<br>
<br>
</b></font></u><font face="Arial">In 99% of cases, the VB programmer using the
WebBrowser control writes the data to a temp file and then loads that file into
the control. Now that we know about the Design Mode, we can write data directly
to the control.</font></p>
<p><b><i><font face="Arial" size="2">WebBrowser1.Document.Script.Document.Clear<br>
WebBrowser1.Document.Script.Document.Write "Hello World"<br>
WebBrowser1.Document.Script.Document.Close</font></i></b></p>
<p><font face="Arial">If we want to get the data back we use a rather obscure
method:</font></p>
<p><i><b><font face="Arial" size="2">strHTMLData = WebBrowser1.Document.All(0).InnerHTML</font></b></i></p>
<p> </p>
<p><u><b><font face="Arial" size="4">Conclusion</font></b></u></p>
<p><font face="Arial">It was not my intention to write a definitive guide on how
to do this but rather to provide an overview and encourage other programmers to
take this further. This is a very powerful feature that is so simple to use yet very little is available in VB.</font></p>
<p> </p>
<p> </p>
</body>

