<div align="center">

## Internet Programming


</div>

### Description

Visit us at http://www.vbparadise.com. In this tutorial I will cover how to perform file transfers between your PC and a web server. The topics of web site management, dynamic generation of web pages, and control/script inclusion in web pages is covered in a different tutorial.

In the Professional and Enterprise Editions of VB are the two controls which provide Internet features - the Internet Transfer Control and the Web Browser Control. The Internet Transfer Control (ITC) provides both FTP and HTTP file transfer controls and is the subject of this tutorial - as is the use of the wininet.dll, which also provides file transfer capabilities that you can access from within your VB applications.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Abdulaziz Alfoudari](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/abdulaziz-alfoudari.md)
**Level**          |Advanced
**User Rating**    |4.2 (50 globes from 12 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Internet/ HTML](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/internet-html__1-34.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/abdulaziz-alfoudari-internet-programming__1-31310/archive/master.zip)





### Source Code

<h4 style="margin-bottom: 5" align="center"><font face="Verdana" size="5">Visit
Us At: <font color="#800000"><span style="font-weight: 400">
<a href="http://www.vbparadise.com">http://www.vbparadise.com</a></span></font></font></h4>
<h4 style="margin-bottom: 5">&nbsp;</h4>
<h4 style="margin-bottom: 5"><font face="Verdana" size="2">Internet - File
Transfers</font></h4>
<p style="margin-top: 5"><font face="Verdana" size="2">In this tutorial I will
cover how to perform file transfers between your PC and a web server. The topics
of web site management, dynamic generation of web pages, and control/script
inclusion in web pages is covered in a different tutorial. </font><p>
<font face="Verdana" size="2">In the Professional and Enterprise Editions of VB
are the two controls which provide Internet features - the Internet Transfer
Control and the Web Browser Control. The Internet Transfer Control (ITC)
provides both FTP and HTTP file transfer controls and is the subject of this
tutorial - as is the use of the wininet.dll, which also provides file transfer
capabilities that you can access from within your VB applications. </font><p>
<font face="Verdana" size="2">The wininet.dll file is simply a library of
functions that you can use to do file transfers. It is <b>not</b> a part of VB6,
but <b>is</b> installed along with the Microsoft Internet Explorer. This means
you can safely assume that it is available on most PCs. </font><p>
<font face="Verdana" size="2">The reason for using the ITC is simplicity. The
control provides a very straightforward, fairly simple interface to use in your
VB programs. The down side is that the control is about 126K in size, increasing
the size of the installation files your application users will have to download.
</font><p><font face="Verdana" size="2">By using the existing wininet.dll file,
you eliminate the increased application distribution file size, plus the
wininet.dll offers greater control over the file transfer process. As you would
expect, the wininet.dll is harder to learn and is not documented as part of the
VB documentation package. </font><p><font face="Verdana" size="2">So, I'll focus
first on the ITC, then end with enough detail on wininet.dll to show you how to
use it as an alternative to the ITC. </font><p>
<h4 style="margin-bottom: 5"><font face="Verdana" size="2">Protocols</font></h4>
<p style="margin-top: 5"><font face="Verdana" size="2">For our purposes, only
two of the many protocols matter - HTTP and FTP. Protocols are simply rules
which programmers have agreed upon, and documented so that anyone using the
protocol can be assured that their software can communicate with other programs
which also use that protocol. </font><p><font face="Verdana" size="2">For the
purposes of file transfer as discussed in this tutorial, It's not really that
important to understand the details of either protocol, but there are a few
facts which are important to know. The key point to remember right now is that
the code for using the ITC for an FTP or for an HTTP file transfer are slightly
different. </font><p><font face="Verdana" size="2">Most web sites are accessible
by both FTP and HTTP (i.e., servers typically run both FTP and HTTP server
software for accessing content), so you can usually chose which approach to
take. </font><p><font face="Verdana" size="2">In general, it really makes little
difference which protocol you use. Many of the books I've read recommend the FTP
protocol because it is more flexible than HTTP (i.e., read that as FTP has more
features which can be controlled by code than does HTTP). I generally agree with
that recommendation, but will confess that my own freeware applications now have
an online update feature which is based on the HTTP protocol. </font><p>
<font face="Verdana" size="2">Here's the tradeoff that drove me to that
decision: </font>
<ul>
       <li><font face="Verdana" size="2">FTP - Many ISPs do not allow
       anonymous (i.e., no username and no password) FTP connections to a
       website. But, I do not want to put my username/password into a
       distributed application for fear of compromising security on my
       web site. I could password protect and expose just a particular
       directory on the web site, but I've chosen to take no risks in
       that area. </font></li>
       <li><font face="Verdana" size="2">HTTP - The ITC code for an HTTP
       file transfer is extremly simple (the FTP code is not that
       complicated, it's just that the HTTP code is simpler). </font>
       </li>
</ul>
<p><font face="Verdana" size="2">One of the key drawbacks of using the ITC for
file transfer (regardless of the protocol that is used) is that it does not
provide any built-in capability to identify how many bytes of the transfer are
complete at any point in time. All you can tell from within your VB program is
whether the file transfer is in progress, or that it has stopped (because of a
successful transfer or some error that stopped the file transfer process). This
is one of the key reasons you might be interested in looking at one of the 3rd
party file transfer OCXs. </font><p style="margin-bottom: 5">
<p style="margin-bottom: 5"><font face="Verdana" size="2"><b>OpenURL Method of
File Transfer</b> </font><p style="margin-top: 5"><font face="Verdana" size="2">
Now that I've gotten the introductory comments out of the way, let's talk about
the details of the ITC. The two most important things to know about the ITC is
that there are two methods of downloading files from a web site - the OpenURL
method and the Execute method. Both support the FTP and HTTP file transfer
protocols. </font><p><font face="Verdana" size="2">The <b>OpenURL</b> method is
very simple. You put in a file name to download and tell the program whether the
file is all text or binary. The code looks for an HTTP transfer of a text file
looks like this:<br>
</font>
<pre><font face="Verdana">text1.text = inet1.OpenURL (&quot;http://www.vbinformation.com/badclick.htm&quot;, icString)
</font></pre>
<p><font face="Verdana" size="2">The code for an HTTP transfer of a binary file
looks like this:<br>
</font>
<pre><font face="Verdana">Dim bData() as Byte
bData() = inet1.OpenURL (&quot;http://www.vbinformation.com/badclick.htm&quot;, icByteArray)
</font></pre>
<p><font face="Verdana" size="2">Since all files (text or binary) can be
transferred as a binary file, I used the same file name in both examples. Note
that in the first case, the downloaded file content is placed in a textbox named
'text1'. In the second case, the downloaded file content is saved in a Byte
array whose upper bound is set by the number of bytes downloaded by the OpenURL
method. Also, note that both examples use HTTP URLs, but FTP URLs could have
been used just as readily. </font><p><font face="Verdana" size="2">In case you
don't remember, an easy way to save the bData byte array is:<br>
</font>
<pre><font face="Verdana">Open &quot;filename&quot; for Binary as #1
Put #1, , bData()
Close #1
</font></pre>
<p><font face="Verdana" size="2">This is really all there is to successfully
downloading a file by using the OpenURL method. I'll cover the question of
errors (such as when the server is down, or the file is not there) later in this
tutorial. </font><p><font face="Verdana" size="2">You should note that the
OpenURL method is synchronous - which simply means that any code that follows
the OpenURL statement will not be executed until the file transfer is completed,
or until the file transfer is stopped by the occurence of an error or by a user
command (I'll show how to do this later). </font><p style="margin-bottom: 5">
<p style="margin-bottom: 5"><font face="Verdana" size="2"><b>Execute Method of
File Transfer</b> </font><p style="margin-top: 5"><font face="Verdana" size="2">
The second method for downloading a file is the Execute method. As you'll see it
provides more features, but is definitely more complicated to code. The one key
difference that you'll want to be aware of is that with the Execute method the
bytes of data are sometimes, but not always, kept within the ITC itself (in a
memory buffer). When the ITC does keep the downloaded bytes in its buffer, you
must use another method called GetChunk to extract the dowloaded bytes. Whether
the memory buffer is used varies with the arguments used in calling the Execute
method. I'll give more detail on that later. </font><p>
<font face="Verdana" size="2">Another key difference that you should know about
is the Execute method is asynchronous - meaning that it will download the file
in the background and that any code following the Execute statement will be
executed immediately. </font><p><font face="Verdana" size="2">Finally, to
complicate the discussion a bit more, the arguments you use in the Execute
method differ depending on whether you want to use FTP or HTTP for the file
transfer. </font><p><font face="Verdana" size="2">Here's the general syntax for
the Execute method:<br>
</font>
<pre><font face="Verdana">inet1.Execute (url, operation, data, requestheaders)
</font></pre>
<p><font face="Verdana" size="2">For an FTP file transfer, only the first two
arguments are used. For an HTTP file transfer, all four arguments may be used.
</font><p><font face="Verdana" size="2">Here's an example of the code you would
use to start a transfer using the Execute method and the FTP protocol:<br>
</font>
<pre><font face="Verdana">inet1.Execute (&quot;ftp://www.microsoft.com&quot;, &quot;DIR&quot;)
</font></pre>
<p><font face="Verdana" size="2">This command transfers the directory listing of
the Microsoft ftp site. Note than while the OpenURL method returns data to a
variable or an array, the Execute method does not! The data returned by the
Execute method will either be kept within the ITC's buffer, or be directed to a
file according to the specifics of the command it is given. </font><p>
<font face="Verdana" size="2">The Execute method actually supports 14 FTP
commands (which are placed in the 'operation' argument), but there are primarily
three (CD, GET, and PUT) which you will use most often:<br>
</font>
<ul>
       <li><font face="Verdana" size="2">inet1.Execute
       (&quot;ftp://www.microsoft.com&quot;, &quot;CD newdirectory&quot; </font></li>
       <li><font face="Verdana" size="2">inet1.Execute
       (&quot;ftp://www.microsoft.com&quot;, &quot;GET remotefile localfile&quot; </font>
       </li>
       <li><font face="Verdana" size="2">inet1.Execute
       (&quot;ftp://www.microsoft.com&quot;, &quot;PUT localfile remotefile&quot; </font>
       </li>
</ul>
<p><font face="Verdana" size="2">The first of these three allow you to make the
connection to the FTP server and to navigate to the directory where the files
are located. The second shows how to GET a file from the server and put it on
your PC, while the third shows how to PUT a local file from your PC onto the FTP
server (in the directory to which you navigated to using the CD command). </font>
<p><font face="Verdana" size="2">Also, you will note that the GET and PUT
commands create a file on either the local or remote computers. In these cases,
the ITC memory buffer is not used. However, the ITC memory buffer must be
accessed in order to get the output of the 'DIR' command. </font><p>
<font face="Verdana" size="2">In order to discuss how the ITC memory buffer is
accessed, we have to talk first about the StateChanged Event. Statechanged is
the only event the ITC control has, and it provides a variable called 'State'
which must be read to determine the status of a pending Execute method. </font>
<p><font face="Verdana" size="2">The State values are:<br>
</font>
<ul>
       <li><font face="Verdana" size="2">icNone (0) </font></li>
       <li><font face="Verdana" size="2">icHostResolvingHost (1) </font>
       </li>
       <li><font face="Verdana" size="2">icHostResolved (2) </font></li>
       <li><font face="Verdana" size="2">icConnecting (3) </font></li>
       <li><font face="Verdana" size="2">icConnected (4) </font></li>
       <li><font face="Verdana" size="2">icRequesting (5) </font></li>
       <li><font face="Verdana" size="2">icRequestSent (6) </font></li>
       <li><font face="Verdana" size="2">icReceivingResponse (7) </font>
       </li>
       <li><font face="Verdana" size="2">icResponseReceived (8) </font>
       </li>
       <li><font face="Verdana" size="2">icDisconnecting (9) </font></li>
       <li><font face="Verdana" size="2">icDisconnected (10) </font></li>
       <li><font face="Verdana" size="2">icError (11) </font></li>
       <li><font face="Verdana" size="2">icResponseCompleted (12) </font>
       </li>
</ul>
<p><font face="Verdana" size="2">Typically, Select Case code is used within the
StateChanged event to determine what action to take. In general, only actions 8,
11, and 12 are used to generate a code response. The others are used mostly to
decide what message to display in a status label/toolbar. </font><p>
<font face="Verdana" size="2">For State=12, where the file transfer is complete,
the action you take is entirely up to you. This would usually be a simple popup
message telling the user that the file transfer is complete. </font><p>
<font face="Verdana" size="2">For State=11, which indicates that an error has
occurred, you would have to generate code necessary to correct or ignore the
error condition. </font><p><font face="Verdana" size="2">Generally, you simply
wait for State=12 to indicate that the transfer is complete. But, in some cases
you may want to begin extracting data before the transfer is complete. For
example, the HTTP header information is received first, but is not included in
the ITC download buffer. To get that information you use the .GetHeader method.
You can use the State=8 to determine when the header information is available.
</font><p><font face="Verdana" size="2">In those cases where the ITC buffer is
used to temporarily store the downloaded information, the .GetChunk method is
used. Here's the code for the case where string data is being downloaded and a
State=12 has been received to indicate that the transfer is complete:<br>
</font>
<pre><font face="Verdana">Do
 DoEvents
 bData = inet1.GetChunk (1024, icString)
 AllData = AllData &amp; bData
Loop Until bData = &quot;&quot;
</font></pre>
<p><font face="Verdana" size="2">In the case where a State=8 has been received,
it is possible that no actual data is in the ITC buffer (such as when only
header information has been received). So, if the above code is used following a
State=8 event, the condition of bData=&quot;&quot; may not indicate completion of the data
transfer. </font><p><font face="Verdana" size="2">Finally, remember that you may
have to set the .UserName and .Password properties if you are using the FTP
commands within the Execute method. If the FTP site you are accessing allows
'anonymous' logon's, then you will not have to set these properties. </font><p>
<h4 style="margin-bottom: 5"><font face="Verdana" size="2">
Properties/Methods/Events</font></h4>
<p style="margin-top: 5"><font face="Verdana" size="2">As you saw in the sample
code, the basics of file transfer don't require knowledge of all 19 properties,
5 methods, and one event exposed by the ITC. The following table list all of the
ITC interface elements, but as you will see in the discussion that follows, you
will not use but a few of these in most applications: </font><p>
<ul>
       <table>
        <tr>
        <th><font face="Verdana" size="2">Properties </font></th>
        <th><font face="Verdana" size="2">Methods </font></th>
        <th><font face="Verdana" size="2">Events </font></th>
        </tr>
        <tr>
        <td><font face="Verdana" size="2">AccessType </font></td>
        <td><font face="Verdana" size="2">Cancel </font></td>
        <td><font face="Verdana" size="2">StateChanged </font></td>
        </tr>
        <tr>
        <td><font face="Verdana" size="2">Document </font></td>
        <td><font face="Verdana" size="2">Execute </font></td>
        <td></td>
        </tr>
        <tr>
        <td><font face="Verdana" size="2">hInternet </font></td>
        <td><font face="Verdana" size="2">GetChunk </font></td>
        <td></td>
        </tr>
        <tr>
        <td><font face="Verdana" size="2">Password </font></td>
        <td><font face="Verdana" size="2">GetHeader </font></td>
        <td></td>
        </tr>
        <tr>
        <td><font face="Verdana" size="2">Protocol </font></td>
        <td><font face="Verdana" size="2">OpenURL </font></td>
        <td></td>
        </tr>
        <tr>
        <td><font face="Verdana" size="2">Proxy </font></td>
        <td></td>
        <td></td>
        </tr>
        <tr>
        <td><font face="Verdana" size="2">RemoteHost </font></td>
        <td></td>
        <td></td>
        </tr>
        <tr>
        <td><font face="Verdana" size="2">RequestTimeout </font></td>
        <td></td>
        <td></td>
        </tr>
        <tr>
        <td><font face="Verdana" size="2">ResponseCode </font></td>
        <td></td>
        <td></td>
        </tr>
        <tr>
        <td><font face="Verdana" size="2">ResponseInfo </font></td>
        <td></td>
        <td></td>
        </tr>
        <tr>
        <td><font face="Verdana" size="2">StillExcuting </font></td>
        <td></td>
        <td></td>
        </tr>
        <tr>
        <td><font face="Verdana" size="2">URL </font></td>
        <td></td>
        <td></td>
        </tr>
        <tr>
        <td><font face="Verdana" size="2">UserName </font></td>
        <td></td>
        <td></td>
        </tr>
       </table>
</ul>
<p><font face="Verdana" size="2">In addition, the normal control properties of
.Index, .Left, .Top, .Tag, .Parent and .hInternet are available for the ITC.
</font><p><font face="Verdana" size="2">This table can be digested more easily
if you think in terms of how the properties/methods/events are used. Here's the
way I've grouped them in my notes: </font>
<ul>
       <li><font face="Verdana" size="2"><b>Basic</b><br>
       The .OpenURL and .Execute methods are the heart of using the ITC.
       Everything you do requires the use of one of these two methods.
       The .GetChunk method is used to capture data downloaded by the
       .Execute method. </font></li>
       <li><font face="Verdana" size="2"><b>Working</b><br>
       The .URL, .Cancel, .StillExecuting, .ResponseCode, .ResponseInfo,
       .Cancel, .GetChunk, and .GetHeader interface elements are used
       extensively during program execution. </font></li>
       <li><font face="Verdana" size="2"><b>Startup</b><br>
       The .AccessType, .Proxy, .Protocol, .RequestTimeout, .RemoteHost,
       .UserName, .Password, and .RemotePort are very basic properties
       which you set once or use the default - then you're done with
       them. </font></li>
</ul>
<p><font face="Verdana" size="2">The bottom line of this tutorial section is
that file transfers can be made very easily using the ITC and just a handful of
the properties, methods and events supported by the control. </font><p>
<font face="Verdana" size="2">As a final note, in case you've been watching
carefully you will notice that I left out any discussion on the use of the
.Execute method with HTTP commands. This was strictly for lack of time/space in
this tutorial. The .Execute method can be used equally with either FTP or HTTP
commands, but the FTP options are generally more extensive so FTP is the normal
choice for programmers. </font>
<h4><font face="Verdana" size="2">WinInet.dll - File Transfer Alternative</font></h4>
<p><font face="Verdana" size="2">--- info on using wininet.dll goes here ----</font>

