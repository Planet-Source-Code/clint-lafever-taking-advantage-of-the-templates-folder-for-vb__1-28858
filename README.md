<div align="center">

## Taking advantage of the Templates folder for VB


</div>

### Description

Article to show those who do not know about the Templates folder for VB how to save templates to it for future use.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Clint LaFever](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/clint-lafever.md)
**Level**          |Beginner
**User Rating**    |5.0 (50 globes from 10 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Miscellaneous](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/miscellaneous__1-1.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/clint-lafever-taking-advantage-of-the-templates-folder-for-vb__1-28858/archive/master.zip)





### Source Code

<p><font face="Verdana" color="#800000"><b>Taking advantage of the Templates
folder for VB</b></font></p>
<p><small><font face="Verdana">In newsgroups I have seen the question asked a
lot about how to change the default properties of the form when you add new
forms to a project.&nbsp; With how much this question is asked I figured if they
just did a search in the newsgroup they would find the answer without having to
ask again (but that is a different story).&nbsp; Anyhow, I decided to post this
here to help all the newbies out there and so I can use this link in my
responses to the newsgroups questions.</font></small></p>
<p><small><font face="Verdana">So, how many times have you started a new project
and found yourself putting in the same old common code you always use in every
project.&nbsp; Some people will just point to a common location that they saved
this code to, others will use an Add-In that stores reusable code to insert it,
and so forth.&nbsp; For the most part all of us have some set of code that we
always want and need in every project.&nbsp; So others always work with
databases and always need to reference ADO, DAO etc.&nbsp; Others have a preferred
font setting for all forms.&nbsp; Well, the simple way to deal with this is to
take advantage of the Templates folder found in where you installed VB.&nbsp;
For me that is D:\Program Files\Microsoft Visual Studio\VB98\Template.&nbsp; If
you go to your folder location you will see this folder contains sub folders for
templates like:</font></small></p>
<ul>
 <li><small><font face="Verdana" color="#000080">Classes</font></small></li>
 <li><small><font face="Verdana" color="#000080">Code</font></small></li>
 <li><small><font face="Verdana" color="#000080">Controls</font></small></li>
 <li><small><font face="Verdana" color="#000080">Forms</font></small></li>
 <li><small><font face="Verdana" color="#000080">MDIForms</font></small></li>
 <li><small><font face="Verdana" color="#000080">Menus</font></small></li>
 <li><small><font face="Verdana" color="#000080">Projects</font></small></li>
 <li><small><font face="Verdana" color="#000080">Proppage</font></small></li>
 <li><small><font face="Verdana" color="#000080">Userctls</font></small></li>
 <li><small><font face="Verdana" color="#000080">Userdocs</font></small></li>
</ul>
<p><small><font face="Verdana">Now the smart ones out there who never saw this
before may be catching on already.&nbsp; Ok, now let me show you how to use this
for making a Project Template.</font></small></p>
<ol>
 <li><small><font face="Verdana" color="#000080">Open VB.&nbsp; Start a new
  standard exe project.</font></small></li>
 <li><small><font face="Verdana" color="#000080">Add all the modules, classes,
  forms, references, components, etc that you need</font></small></li>
 <li><small><font face="Verdana" color="#000080">You may even consider setting
  some project properties like Copyright etc.</font></small></li>
 <li><small><font face="Verdana" color="#000080">Make sure you have all your
  modules and forms good meaningful names as to not overwrite any other files
  later (you will see).&nbsp; As for the Project Name, save it with a nice English
  like file name.</font></small></li>
 <li><small><font face="Verdana" color="#000080">Now save ALL these files to D:\Program Files\Microsoft Visual Studio\VB98\Template\Projects
  (<b>note to use your path not mine</b>)</font></small></li>
 <li><small><font face="Verdana" color="#000080">Close VB.</font></small></li>
 <li><small><font face="Verdana" color="#000080">Open VB.</font></small></li>
</ol>
<p><small><font face="Verdana">Now you should see that project as an option of a
template for starting a new project.&nbsp; Choose it to start your new project
and presto, you have ALL your code, and property settings all in place.&nbsp;
Easy huh.</font></small></p>
<ol>
 <li><small><font face="Verdana" color="#000080">Now, go ahead and start just a
  standard EXE project.&nbsp;&nbsp;</font></small></li>
 <li><small><font face="Verdana" color="#000080">On the that basic first form,
  set it up with all the properties you like using for all your forms.&nbsp;
  Now save </font><font face="Verdana" color="#800000"><b>JUST THAT FORM</b></font><font face="Verdana" color="#000080">
  to D:\Program Files\Microsoft Visual Studio\VB98\Template\Forms (<b>remember,
  use meaningful names and not to overwrite existing templates</b>.)</font></small></li>
 <li><small><font face="Verdana" color="#000080">Close VB.</font></small></li>
 <li><small><font face="Verdana" color="#000080">Open VB.</font></small></li>
 <li><small><font face="Verdana" color="#000080">Start a new project form you
  nice new template.</font></small></li>
 <li><small><font face="Verdana" color="#000080">Click your toolbar to add a
  new form.</font></small></li>
</ol>
<p><small><font face="Verdana">Your new template of a form is now an option of
one to add.</font></small></p>
<p><font face="Verdana"><small>Why settle for the default when you can have it
your way.&nbsp; I think you can see what you can do now, if not, think about
another profession.&nbsp; Just kidding.&nbsp; I hope this helps everyone who did
not know about this.&nbsp; I just find it so easy to be able to start a new
project and have all the references, components, and code I always use and need
already there.</small></font></p>
<p><font face="Verdana"><small>-Clint <a href="mailto:LaFeverlafeverc@hotmail.com">LaFever<br>
lafeverc@hotmail.com</a></small></font></p>
<p><font face="Verdana"><small><a href="http://vbasic.iscool.net">http://vbasic.iscool.net</a></small></font></p>
<p><font face="Verdana"><small><a href="mailto:LaFeverlafeverc@hotmail.com"><br>
</a></small></font></p>
<p>&nbsp;</p>
<p>&nbsp;</p>
<p>&nbsp;</p>
<p>&nbsp;</p>

