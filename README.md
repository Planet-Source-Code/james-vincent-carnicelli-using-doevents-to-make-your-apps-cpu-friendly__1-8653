<div align="center">

## Using DoEvents to Make Your Apps CPU\-Friendly


</div>

### Description

How to keep your VB application from being a CPU hog using DoEvents.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[James Vincent Carnicelli](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/james-vincent-carnicelli.md)
**Level**          |Beginner
**User Rating**    |4.3 (26 globes from 6 users)
**Compatibility**  |VB 3\.0, VB 4\.0 \(16\-bit\), VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Windows System Services](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/windows-system-services__1-35.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/james-vincent-carnicelli-using-doevents-to-make-your-apps-cpu-friendly__1-8653/archive/master.zip)





### Source Code

If you are a new VB programmer and have begun to develop more sophisticated applications to do animations, heavy number crunching, etc., you may have noticed that sometimes those apps seem to take control of Windows while they run. For example, your mouse clicks and keystrokes may take a long time to register with your app and others. You may not even see your mouse moving with you fast enough.
<P>The good news is that it's fairly easy to correct this problem.
<P>While newer versions of Windows support "simultaneous" multitasking using "time slices" for each process, Windows is still a non-preemptive operating system at its core. "Preemptive multitasking" means that the operating system (OS) gives each running process a slice of time running on the CPU before it interrupts it to give CPU control to the next process. Each process need not care about the CPU needs of any other. "Nonpreemptive multitasking" means that the processes are expected to voluntarily yield control of the CPU to the OS so it can give control to the next running program.
<P>Roughly speaking, if you're not somehow giving control of the CPU back to Windows, other apps can't use it.
<P>Most simple VB programs get control when Windows triggers an event, like a button click or mouse movement. If your app responds to the event, it automatically gives control back to Windows when the responding event (e.g., <TT>Command1_Click()</TT> ) is done executing. But if it doesn't exit within a few seconds, you may start to notice your app is hogging resources.
<P>Fortunately, VB comes with a built in routine to voluntarily give control of the CPU back to Windows for a while: DoEvents. Consider the following simple program:
<UL><PRE>
<FONT COLOR="#000066">Private</FONT> GoForIt <FONT COLOR="#000066">As Boolean</FONT>
<P><FONT COLOR="#000066">Private Sub</FONT> Command1_Click()
&nbsp;&nbsp;&nbsp;&nbsp;<FONT COLOR="#000066">If</FONT> GoForIt <FONT COLOR="#000066">Then</FONT> <FONT COLOR="#006600">'Clicked to stop</FONT>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;GoForIt = <FONT COLOR="#000066">False</FONT>
&nbsp;&nbsp;&nbsp;&nbsp;Else <FONT COLOR="#006600">'Clicked to start</FONT>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;GoForIt = <FONT COLOR="#000066">True</FONT>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<FONT COLOR="#000066">Do While</FONT> GoForIt
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Command1.Caption = Rnd
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<FONT COLOR="#CC0000"><B>DoEvents</B></FONT>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<FONT COLOR="#000066">Loop</FONT>
&nbsp;&nbsp;&nbsp;&nbsp;<FONT COLOR="#000066">End If</FONT>
<FONT COLOR="#000066">End Sub</FONT>
<P><FONT COLOR="#000066">Private Sub</FONT> Form_Unload(Cancel As <FONT COLOR="#000066">Integer</FONT>)
&nbsp;&nbsp;&nbsp;&nbsp;GoForIt = <FONT COLOR="#000066">False</FONT> <FONT COLOR="#006600">'Break out of the loop</FONT>
&nbsp;&nbsp;&nbsp;&nbsp;<FONT COLOR="#CC0000"><B>DoEvents</B></FONT>
<FONT COLOR="#000066">End Sub</FONT>
</PRE></UL>
<P>Notice how DoEvents is used in <TT>Form_Unload()</TT> to let the loop initiated in <TT>Command1_Click()</TT>, if it's still running, finish and exit? That's right; one chunk of your own code can be running "in parallel" with another chunk in your program. You don't have to mess with multithreading or multiprocessing to have this happen. This is another benefit of using DoEvents liberally and carefully.
<P>In summary, use DoEvents to break up algorithms that take a long time or loop continuously to give Windows a chance now and then to do other things with the CPU.

