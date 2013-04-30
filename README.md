wpf_from_com
============

Bringing new life into older applications written in VB6 isn't easy, especially if you've never worked with the language. Personally, unless there were no other option, I wouldn't touch VB6 with a six foot pole. Luckily there is a way to extend an older application while using some of the latest and greatest bells and whistles of WPF and .NET.<br />
<br />
In this post I'll go over one approach that happen to come in handy on a recent project. In general, the solution was to create a COM accessible assembly to serve as the entry point in invoking a WPF application. Sounds simple enough, and it is, with the exception of one or to things that might not be so obvious and which could lead a fair bit of head scratching.<br />
<br />
Basic Steps:<br />
<br />
Create a new project solution and add a WPF application project. If needed, you could also convert the WPF project to a class library. If you do chose to make the project a .DLL, you will also need to delete the App.xaml and code behind files. That's it for the WPF project. Obviously you'd devote more time to the development of your UI but this works for our purposes.<br />
<ul>
</ul>
Next, add a class library project to the solution. There will be two settings that need to be changed.<br />
<ol><ul>
<li>Make the assembly COM accessible: <i>Properties -&gt; Application -&gt; Assembly Information.</i><div class="separator" style="clear: both; text-align: center;">
<a href="http://1.bp.blogspot.com/-0rCN5YJLF9E/UWzFDETjKlI/AAAAAAAAAmg/JNVlvQqt_r8/s1600/makeComVisible.png" imageanchor="1" style="margin-left: 1em; margin-right: 1em;"><img border="0" src="http://1.bp.blogspot.com/-0rCN5YJLF9E/UWzFDETjKlI/AAAAAAAAAmg/JNVlvQqt_r8/s1600/makeComVisible.png" height="167" width="400" /></a></div>
</li>
</ul>
</ol>
<div>
<ol><ul>
<li>Register the assembly for COM interOp. This will register your assembly when you build it and make it accessible through COM. <i>Properties -&gt; Build&nbsp;</i></li>
</ul>
</ol>
<div class="separator" style="clear: both; text-align: center;">
<a href="http://1.bp.blogspot.com/-Lo0TTPO2kOM/UWzFyO3dWJI/AAAAAAAAAmo/ozljNI1Z5wA/s1600/registerForComInterop.png" imageanchor="1" style="margin-left: 1em; margin-right: 1em;"><img border="0" src="http://1.bp.blogspot.com/-Lo0TTPO2kOM/UWzFyO3dWJI/AAAAAAAAAmo/ozljNI1Z5wA/s1600/registerForComInterop.png" height="151" width="400" /></a></div>
<div class="separator" style="clear: both; text-align: center;">
</div>
Declare a COM interface which looks just like a regular interface with the addition of a Guid attribute.<br />
<div class="separator" style="clear: both; text-align: center;">
<a href="http://3.bp.blogspot.com/-QZfXZIrpjco/UWzMzxdahlI/AAAAAAAAAm0/JjZOIW-UpMc/s1600/COM_Interface.png" imageanchor="1" style="margin-left: 1em; margin-right: 1em;"><img border="0" src="http://3.bp.blogspot.com/-QZfXZIrpjco/UWzMzxdahlI/AAAAAAAAAm0/JjZOIW-UpMc/s1600/COM_Interface.png" height="113" width="400" /></a></div>
<br />
<ul>
</ul>
The next step is to create a class that implements the previously defined COM interface. There is nothing special about this class other than the use of the three InteropServies attributes Guid, ClassInterface, and ProgId.<br />
<div class="separator" style="clear: both; text-align: center;">
<a href="http://1.bp.blogspot.com/-GXaktS58ROE/UWzQ65RCBqI/AAAAAAAAAnA/XoxrlwaEAYs/s1600/WPF_Invoker.png" imageanchor="1" style="margin-left: 1em; margin-right: 1em;"><img border="0" src="http://1.bp.blogspot.com/-GXaktS58ROE/UWzQ65RCBqI/AAAAAAAAAnA/XoxrlwaEAYs/s1600/WPF_Invoker.png" height="186" width="400" /></a></div>
<ul>
</ul>
<div style="text-align: left;">
In the ShowWindow method implementation you need to setup an Application instance. There may be only one Application instance per AppDomain. You then set its MainWindow attribute to a window defined in the previously created WPF project. Then you invoke the window by calling "ShowDialog()" off of the applications MainWindow property.<br />
<div class="separator" style="clear: both; text-align: center;">
<a href="http://2.bp.blogspot.com/-vavPgKnE03I/UWzTzIAvNnI/AAAAAAAAAnM/OOiITYfyTeY/s1600/ShowWindowImpl.png" imageanchor="1" style="margin-left: 1em; margin-right: 1em;"><img border="0" src="http://2.bp.blogspot.com/-vavPgKnE03I/UWzTzIAvNnI/AAAAAAAAAnM/OOiITYfyTeY/s1600/ShowWindowImpl.png" /></a></div>
<div class="separator" style="clear: both; text-align: left;">
<br /></div>
<div class="separator" style="clear: both; text-align: left;">
Another way to bring up the UI could be to invoke the Show() method off off the main window followed by the Run() method off of the application instance.&nbsp;</div>
<div class="separator" style="clear: both; text-align: center;">
<a href="http://3.bp.blogspot.com/-7BpDphfxJr8/UWzgL0RfKLI/AAAAAAAAAnc/mkSNTDuqBBs/s1600/New+Tab.png" imageanchor="1" style="margin-left: 1em; margin-right: 1em;"><img border="0" src="http://3.bp.blogspot.com/-7BpDphfxJr8/UWzgL0RfKLI/AAAAAAAAAnc/mkSNTDuqBBs/s1600/New+Tab.png" height="103" width="400" /></a></div>
</div>
<div style="text-align: left;">
<br /></div>
<div style="text-align: left;">
In general, even though you can invoke the UI without creating an Application instance, you should always setup the application for the AppDomain. Especially if you plan to spawn off any worker threads that will at some point need to sync up&nbsp;with UI thread through the Dispatcher. If no application is setup you will receive a null pointer exception when calling Application.Current.Dispatcher.<br />
<br />
<br />
The next item needed is some extra logic to compensate for how/where your COM accessible assembly is called.<br />
<br />
<div class="separator" style="clear: both; text-align: center;">
<a href="http://3.bp.blogspot.com/-VeJynXMshdw/UWzjuY_3PLI/AAAAAAAAAns/I_7dtP8jTmc/s1600/wpfInvokerConstructor.png" imageanchor="1" style="margin-left: 1em; margin-right: 1em;"><img border="0" src="http://3.bp.blogspot.com/-VeJynXMshdw/UWzjuY_3PLI/AAAAAAAAAns/I_7dtP8jTmc/s1600/wpfInvokerConstructor.png" /></a></div>
<br /></div>
<div style="text-align: left;">
<br /></div>
<div style="text-align: left;">
At this point you now have a solution with two projects. One is a WPF application or assembly, the other is a COM accessible class library. So, how would you go about running this? Simple, you just write yourself a small VBScript to test it out. This script uses Wscript.CreateObject("ProgId") to create and return a reference to our COM object. After getting a reference to our object it proceeds to invoke the "ShowWindow()" method defined in our interface.<br />
<br />
<div class="separator" style="clear: both; text-align: center;">
<a href="http://4.bp.blogspot.com/-PNPPXKet0cQ/UWzmCaT5mXI/AAAAAAAAAn8/YDbvT0RkVYU/s1600/showWindowVbs.png" imageanchor="1" style="margin-left: 1em; margin-right: 1em;"><img border="0" src="http://4.bp.blogspot.com/-PNPPXKet0cQ/UWzmCaT5mXI/AAAAAAAAAn8/YDbvT0RkVYU/s1600/showWindowVbs.png" /></a></div>
</div>
<div style="text-align: left;">
<br />
You can then call the above script from the command prompt using CScript.exe:<br />
<div style="text-align: left;">
<i>- &nbsp; C:\Windows\SysWow64\cscript.exe ShowWindow.vbs</i></div>
<div style="text-align: left;">
<i><br /></i></div>
<div style="text-align: left;">
The above script and command can also be used for debugging through Visual Studio. You just have to start CScript.exe as the external program the debugger attaches too.&nbsp;</div>
<div style="text-align: left;">
<br /></div>
<div class="separator" style="clear: both; text-align: center;">
<a href="http://2.bp.blogspot.com/-sT_gTMtRlWg/UW4Y1wgqZwI/AAAAAAAAAoM/Ey2VTsNZugI/s1600/wpfInvokerDebug.png" imageanchor="1" style="margin-left: 1em; margin-right: 1em;"><img border="0" src="http://2.bp.blogspot.com/-sT_gTMtRlWg/UW4Y1wgqZwI/AAAAAAAAAoM/Ey2VTsNZugI/s1600/wpfInvokerDebug.png" /></a></div>
<div class="separator" style="clear: both; text-align: center;">
<br /></div>
<div class="separator" style="clear: both; text-align: center;">
<br /></div>
</div>
<div style="text-align: left;">
<br /></div>
<div style="text-align: left;">
<br /></div>
<div style="text-align: left;">
<br /></div>
<div style="text-align: left;">
<br /></div>
<div style="text-align: left;">
<br /></div>
</div>
