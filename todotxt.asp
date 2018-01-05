<!-- #include file="config.inc" --> 

<head>

<LINK REL="SHORTCUT ICON" HREF="check.ico">

<SCRIPT LANGUAGE="JavaScript" SRC="CalendarPopup.js"></SCRIPT>

<SCRIPT LANGUAGE="JavaScript">document.write(getCalendarStyles());</SCRIPT>


<title>ToDo.ASP</title>
<style>

td
	{
	font-size:small;
	vertical-align:top;
	}
	
<%
if todotxtTheme="DARK" then
	link = "white"
	hover = "yellow"
	background = "black"
	CalBack = "darkblue"
else
	link = "black"
	hover = "green"
	background = "white"
end if
%>
body {
	font-size:small;
	font-color:<%response.write(link)%>;
	background-color:<%response.write(background)%>;
	color:<%response.write(link)%>;
	}
a:link {color:<%response.write(link)%>;}
a:visited {color:<%response.write(link)%>;}
a:hover {color:<%response.write(hover)%>;}
a:active {color:<%response.write(link)%>;}
a:link {text-decoration:none;}
a:visited {text-decoration:none;}
a:hover {text-decoration:none;}
a:active {text-decoration:none;}

.td-nowrap
	{
	white-space:nowrap;
	}
	
	.TESTcpYearNavigation,
	.TESTcpMonthNavigation
			{
			background-color:#6677DD;
			text-align:center;
			vertical-align:center;
			text-decoration:none;
			color:#FFFFFF;
			font-weight:bold;
			}
	.TESTcpDayColumnHeader,
	.TESTcpYearNavigation,
	.TESTcpMonthNavigation,
	.TESTcpCurrentMonthDate,
	.TESTcpCurrentMonthDateDisabled,
	.TESTcpOtherMonthDate
			{
			background-color:<%response.write(CalBack)%>;
			}
	.TESTcpOtherMonthDateDisabled,
	.TESTcpCurrentDate,
	.TESTcpCurrentDateDisabled,
	.TESTcpTodayText,
	.TESTcpTodayTextDisabled,
	.TESTcpText
			{
			font-family:arial;
			font-size:8pt;
			}
	TD.TESTcpDayColumnHeader
			{
			text-align:right;
			border:solid thin #6677DD;
			border-width:0 0 1 0;
			}
	.TESTcpCurrentMonthDate,
	.TESTcpOtherMonthDate,
	.TESTcpCurrentDate
			{
			text-align:right;
			text-decoration:none;
			}
	.TESTcpCurrentMonthDateDisabled,
	.TESTcpOtherMonthDateDisabled,
	.TESTcpCurrentDateDisabled
			{
			color:#D0D0D0;
			text-align:right;
			text-decoration:line-through;
			}
	.TESTcpCurrentMonthDate
			{
			color:#6677DD;
			font-weight:bold;
			}
	.TESTcpCurrentDate
			{
			color: #FFFFFF;
			font-weight:bold;
			}
	.TESTcpOtherMonthDate
			{
			color:#808080;
			}
	TD.TESTcpCurrentDate
			{
			color:#FFFFFF;
			background-color: #6677DD;
			border-width:1;
			border:solid thin #000000;
			}
	TD.TESTcpCurrentDateDisabled
			{
			border-width:1;
			border:solid thin #FFAAAA;
			}
	TD.TESTcpTodayText,
	TD.TESTcpTodayTextDisabled
			{
			border:solid thin #6677DD;
			border-width:1 0 0 0;
			}
	A.TESTcpTodayText,
	SPAN.TESTcpTodayTextDisabled
			{
			height:20px;
			}
	A.TESTcpTodayText
			{
			color:#6677DD;
			font-weight:bold;
			}
	SPAN.TESTcpTodayTextDisabled
			{
			color:#D0D0D0;
			}
	.TESTcpBorder
			{
			border:solid thin #6677DD;
			background-color:<%response.write(CalBack)%>;
			}
</style>
</head>
<body OnLoad="document.NewTask.newPri.focus();">



<%



Set objFSO = CreateObject("Scripting.FileSystemObject")

searchstr = request.querystring("searchstr")
if searchstr="" then searchstr = " "
hide_hidden=0
if searchstr="today" then
	searchstr="due:"&year(now())&"-"&right("0"&month(now()),2) & "-"&right("0" & day(now()),2)
    hide_hidden=1
end if
if searchstr="tomorrow" then searchstr="due:"&year(now()+1)&"-"&right("0"&month(now()+1),2) & "-"&right("0" & day(now()+1),2)

rightSearch=right(searchstr,10)

UpdateLine = request.querystring("UpdateLine")
Complete = request.querystring("Complete")
Postpone = request.querystring("postpone")
PPToday = request.querystring("PPToday")
EditLine = request.querystring("EditLine")
Unhide = request.querystring("Unhide")
if request.querystring("Threshold")="Toggle" then
	if Session("Threshold") = True then
		Session("Threshold")=False
	else
		Session("Threshold")=True
	end if
end if

Const ForReading = 1, ForWriting=2, ForAppending=8

' Postpone Today to Tomorrow
If PPToday then

	Set objFSO = CreateObject("Scripting.FileSystemObject")
	Set objFile = objFSO.OpenTextFile(ToDoFile, ForReading)

	strText = objFile.ReadAll
'	objFile.Close
	strNewText = Replace(strText, year(now())&"-"&right("0"&month(now()),2) & "-"&right("0" & day(now()),2), year(now()+1)&"-"&right("0"&month(now()+1),2) & "-"&right("0" & day(now()+1),2))
	strNewText = left (strNewText, len(strNewText)-2)
	Set objFile = objFSO.OpenTextFile(ToDoFile, ForWriting)
	objFile.WriteLine strNewText
	objFile.Close
	searchstr=year(now()+1)&"-"&right("0"&month(now()+1),2) & "-"&right("0" & day(now()+1),2)
End If

' Unhide Hidden tasks
If Unhide then

	Set objFSO = CreateObject("Scripting.FileSystemObject")
	Set objFile = objFSO.OpenTextFile(ToDoFile, ForReading)

	strText = objFile.ReadAll
'	objFile.Close
	strNewText = Replace(strText,"h:1 ","")
	strNewText = left (strNewText, len(strNewText)-2)
	Set objFile = objFSO.OpenTextFile(ToDoFile, ForWriting)
	objFile.WriteLine strNewText
	objFile.Close
End If

ShowTasks=0
TotalTasks = 0


' Archive
if request.querystring("archive")="archive" then
	Set objFile = objFSO.OpenTextFile(ToDoFile, ForReading)
	i = 0
	Dim arrFileLinesArch()
	Do Until objFile.AtEndOfStream
		Redim Preserve arrFileLinesArch(i)
		arrFileLinesArch(i) = objFile.ReadLine
		i = i + 1
	Loop
'	objFile.Close
	Set objFile = objFSO.OpenTextFile(ToDoFile, ForWriting)
	Set objFileArchive = objFSO.OpenTextFile(ArchFile, ForAppending)
	for each StrLine in arrFileLinesArch
		if left(strline,2)="x " then 
			if instr(1,strline,"rec:",1)<1 then
				objFileArchive.writeline (strline)
			end if
		else
			objFile.writeline (strline)
		end if
	next
	objFile.close
	objFileArchive.close
end if



' Add Task
if request.querystring("AddNew")="Add" then
	newLine = request.querystring("newPri") & " " & request.querystring("newStr") & " " & request.querystring("newCat")
	newDue = request.querystring("newDue")
	if newDue<>"None" then
		newDate = request.querystring("customdate")
		if len(newDate)<1 then
			postDays = request.querystring("newDue")
			newDate = (year(now()+postDays)&"-"&right("0"&month(now()+postDays),2) & "-"&right("0" & day(now()+postDays),2))
		end if
		newLine = newLine & " due:" & newDate
	end if
	Set objFile = objFSO.OpenTextFile(ToDoFile, ForAppending)
	objFile.Writeline (newLine)
	objFile.Close
end if



' read tasks
Set objFile = objFSO.OpenTextFile(ToDoFile, ForReading)


Dim arrFileLines()
i = 0
Do Until objFile.AtEndOfStream
	Redim Preserve arrFileLines(i)
	arrFileLines(i) = objFile.ReadLine
	i = i + 1
Loop
objFile.Close

' update tasks
Set objFile = objFSO.OpenTextFile(ToDoFile, ForWriting)
i=0
hide_alert=0
For Each strLine in arrFileLines
	if cstr(UpdateLine)=cstr(strLine) then
		if Complete=1 then
' recur tasks (recurring currently does it from due date, not from date completed, need to change this for using + signs similar to simpletask
' also need to address threshold dates
			if instr(1,arrFileLines(i),"rec:",1) > 0 then
			    CurrDue = mid(arrFileLines(i),instr(1,arrFileLines(i),"due:",1)+4,10)
				RecStr = mid(strline,instr(1,strline,"rec:",1)+4,instr(instr(1,strline,"rec:",1),strline," ",1)-instr(1,strline,"rec:",1)-4)
				if left(RecStr,1)="+" then
					RecStr = right(RecStr,len(RecStr)-1)
					DateRef = CurrDue
				else
					DateRef = now()
				end if
				Select Case right(RecStr, 1)
					Case "d"
						RecDate = dateadd("D",cint(left(RecStr,len(RecStr)-1)), DateRef)
					Case "w"
						RecDate = dateadd("D",cint(left(RecStr,len(RecStr)-1))*7, DateRef)
					Case "m"
						RecDate = dateadd("M",cint(left(RecStr,len(RecStr)-1)), DateRef)
					Case "y"
						RecDate = dateadd("YYYY",cint(left(RecStr,len(RecStr)-1)), DateRef)
				End Select
				strLine = replace (strLine,"h:1 ","")
				strLine = replace (strLine,"due:" & CurrDue,"due:" & Year(RecDate) & "-" & right("0"&Month(RecDate),2) & "-" & right("0"&Day(RecDate),2))
				if instr(1,arrFileLines(i),"t:",1)>0 then
					DateInc = cdate(RecDate) - cdate(CurrDue)
					CurrThresh = mid(arrFileLines(i),instr(1,arrFileLines(i),"t:",1)+2,10)
					ThreshDate = dateadd("D",DateInc,CurrThresh)
					strLine = replace (strLine,"t:" & CurrThresh,"t:" & Year(ThreshDate) & "-" & right("0"&Month(ThreshDate),2) & "-" & right("0"&Day(ThreshDate),2))
				end if
			else
				strLine = replace (strLine,"h:1 ","")
				strLine=("x " & strLine)
			end if
		end if
		if Postpone>0 then
			CurrDue = mid(arrFileLines(i),instr(1,arrFileLines(i),"due:",1)+4,10)
			NewDue = cdate(CurrDue) + Postpone
			NewDue = Year(NewDue) & "-" & right("0"&Month(NewDue),2) & "-" & right("0"&Day(NewDue),2)
			strLine = Replace (strLine,CurrDue,NewDue)
		end if
		if request.querystring("EditLine")="Submit" then strLine=request.querystring("NewLine")

		'change Priority
		PriChange = request.querystring("PriChange")
		if request.querystring("PriChange")="Up" then
			strLine = Replace (strLine,"(B)","(A)")
			strLine = Replace (strLine,"(C)","(B)")
			strLine = Replace (strLine,"(D)","(C)")
			strLine = Replace (strLine,"(E)","(D)")
		end if
		if request.querystring("PriChange")="Down" then
			strLine = Replace (strLine,"(D)","(E)")
			strLine = Replace (strLine,"(C)","(D)")
			strLine = Replace (strLine,"(B)","(C)")
			strLine = Replace (strLine,"(A)","(B)")
		end if
		if (len(cstr(PriChange))>0 and left(strLine,1)<>"(") then strLine = "(C)" & strLine
		
		if request.querystring("Hide")="Hide" then
			if instr(1,strLine,"h:1",1)>1 then
				strLine = Replace (strLine,"h:1 ","")
			else
				strLine = Replace (strLine,"due:","h:1 due:")
			end if
		end if
		arrFileLines(i) = strLine
	end if
	objFile.WriteLine (strLine)
	if instr(1,strline,"h:1",1)>0 then hide_alert=1
	i=i+1
next

if hide_alert=1 then response.write("<b><a href=todotxt.asp?searchstr=h:1><font color=red size=+1>HIDDEN TASKS</a><br></b></font>")
objFile.Close

for a = UBound(arrFileLines) - 1 To 0 Step -1
    for j= 0 to a
	tempJpri = left(arrFileLines(j),3)
	tempJ1pri = left(arrFileLines(j+1),3)
	if instr(1,arrFileLines(j),"due:",1) > 0 then
		tempJdue = mid(arrFileLines(j),instr(1,arrFileLines(j),"due:",1)+4,10)
	else
		tempJdue = "2500-12-31"
	end if

	if instr(1,arrFileLines(j+1),"due:",1) > 0 then
		tempJ1due = mid(arrFileLines(j+1),instr(1,arrFileLines(j+1),"due:",1)+4,10)
	else
		tempJ1due = "2500-12-31"
	end if

	if tempJdue > tempJ1due then
		temp=arrFileLines(j+1)
		arrFileLines(j+1)=arrFileLines(j)
		arrFileLines(j)=temp
	end if

	if tempJdue = tempJ1due then
		if tempJpri > tempJ1pri then
			temp=arrFileLines(j+1)
			arrFileLines(j+1)=arrFileLines(j)
			arrFileLines(j)=temp
		end if
		if tempJpri = tempJ1pri then
			if mid (arrFileLines(j),5,4)>mid(arrFileLines(j+1),5,4) then
				temp=arrFileLines(j+1)
				arrFileLines(j+1)=arrFileLines(j)
				arrFileLines(j)=temp
			end if
		end if
	end if
    next
next
response.write("<table width='100%'><td>")

response.write("<a href=todotxt.asp?searchstr=today>Today</a> | ")

response.write("<a href=todotxt.asp?searchstr=tomorrow>Tomorrow</a> | ")
response.write("<a href=todotxt.asp>All</a> | ")
response.write("<a href=todotxt.asp?searchstr=Completed>Completed</a> | ")
%>
<a href=todotxt.asp?PPToday=True onclick="return confirm('Are you sure?')">PP Today</a> | 
<%
response.write("<a href=todotxt.asp?searchstr="&server.urlencode(request.querystring("searchstr"))&"&	=True>Unhide</a> | ")
if session("Threshold") then response.write ("<b>")
response.write("<a href=todotxt.asp?searchstr="&server.urlencode(request.querystring("searchstr"))&"&Threshold=Toggle>Threshold</a></b>")
response.write(" | <a href=todotxt.asp?searchstr="&server.urlencode(request.querystring("searchstr"))&"&archive=archive>ARCHIVE</a>")
response.write("</td><td align=right>")
response.write(year(now())&"-"&right("0"&month(now()),2)&"-"&right("0"&day(now()),2))
response.write("</td></table>")


response.write("<p>")
if EditLine="1" then
	%><form action=todotxt.asp action=todotxt.asp method=get>
		<input type=text value="<%response.write(UpdateLine)%>" size=100 name=NewLine>
		<input type=hidden value="<%response.write(UpdateLine)%>" name=UpdateLine>
		<input type=hidden value="<%response.write(request.querystring("searchstr"))%>" name=SearchStr>
		<input type=submit name=EditLine value=Submit>
		<input type=submit name=CancelEdit value=Cancel>
	</form>
<% end if %>

<form>
		<input type=text name=searchstr>
		<input type=submit value="Search">
		<%
		response.write("&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;")
		if isdate(rightSearch) then
			response.write("<a href=todotxt.asp?searchstr=" & (year(cdate(rightSearch)-1)&"-"&right("0"&month(cdate(rightSearch)-1),2)&"-"&right("0"&day(cdate(rightSearch)-1),2)) & ">&larr;</a> | ")
		end if
		response.write ("<b>" & searchstr & "</b>")
		if isdate(rightSearch) then
			response.write(" | <a href=todotxt.asp?searchstr=" & (year(cdate(rightSearch)+1)&"-"&right("0"&month(cdate(rightSearch)+1),2)&"-"&right("0"&day(cdate(rightSearch)+1),2)) & ">&rarr;</a>")
		end if%>
</form>



<form name="NewTask" id="NewTask">
	<select name=newPri>
		<option value="(A)">A</option>
		<option value="(B)">B</option>
		<option value="(C)" selected>C</option>
		<option value="(D)">D</option>
		<option value="(E)">E</option>
	</select>
	<input type=text name="newStr">
	<select name=newCat>
	<%
	for each Cat in ArrCats
		response.write("<option value='@"&Cat&"'>"&Cat&"</option>")
	Next
	%>
	</select>
	<select name=newDue>
	<%
	i=0
	do while i<ubound(ArrDueD)
		response.write ("<option value = '"&ArrDueD(i+1)&"'")
		if SelectedDue = ArrDueD(i) then response.write (" selected")
		response.write (">" & ArrDueD(i) & "</option>")
		i=i+2
	loop
	%>
	</select>
	
	<SCRIPT LANGUAGE="JavaScript" ID="js18">
	var cal18 = new CalendarPopup("testdiv1");
	cal18.setCssPrefix("TEST");
	</SCRIPT>
	<INPUT TYPE="text" NAME="CustomDate" VALUE="" onClick="cal18.select(document.NewTask.CustomDate,'anchor18','yyyy-MM-dd'); return false;"NAME="anchor18" ID="anchor18">
	<input type=hidden name=searchstr value=<%response.write (request.querystring("searchstr"))%>>
	<input type=submit name=AddNew value="Add">
</form>


<DIV ID="testdiv1" STYLE="position:absolute;visibility:hidden;background-color:white;layer-background-color:white;"></DIV>
<p>
<font face=verdana size=-1>

<%

' Show Tasks

response.write ("<table border=1 width='100%'>")

For Each strLine in arrFileLines
	TotalTasks=TotalTasks +1
	TaskRec = instr(1,strline,"rec:",1)
	if TaskRec>0 then
		RecStr = mid(strline,TaskRec,instr(TaskRec,strline," ",1)-TaskRec)
	else
		RecStr = ""
	end if

	dueLoc = instr(1,strline,"due:",1)
	if dueLoc > 0 then
		tempdue = mid(strline,dueLoc+4,10)
	else
		tempdue = "2500-12-31"
	end if

	threshLoc = instr(1,strline,"t:",1)
	if threshLoc > 0 then
		tempThresh = mid(strline,threshLoc+2,10)
	else
		tempThresh = "2500-12-31"
	end if
	
	ThreshBarrer = "2500-12-31"
	if isdate(rightSearch) then ThreshBarrier = rightSearch
	
	if (left(strline, 2)<>"x ") and ((instr(1,strline,searchstr,1) > 0) or (left(searchstr,1)="-" and instr(1,strline,right(searchstr,len(searchstr)-1))=0) or cdate(tempdue)<now()-1 or (Session("Threshold")=True and cdate(tempThresh)-1<cdate(ThreshBarrier))) and Not(hide_hidden=1 and instr(1,strline,"h:1",1)>0) then
		ShowTasks = ShowTasks + 1
		if left(strLine,1)="(" then
			tempPri = left(strLine,3)
			tempStr = right(strLine,len(strLine)-3)
		else
			tempStr = strLine
			tempPri = ""
		end if

		TempCat= ""
		for each Cat in ArrCats
			if instr(1,strline,"@"&Cat,1) then TempCat = TempCat & Cat & " "
			tempStr = replace (tempStr,"@" & Cat & " ","")
		Next

		response.write ("<tr><td class=""td-nowrap"" width=1><a href='todotxt.asp?UpdateLine=" & server.urlencode(strLine) & "&SearchStr=" & server.urlencode(request.querystring("searchstr")) & "&Complete=1'><font size=3>&#x25a2;</font></a></td>")
		Select Case tempPri
			case "(A)"
				PriFormIn = PriAForm
				TaskFormIn = TaskAForm
			case "(B)"
				PriFormIn = PriBForm
				TaskFormIn = TaskBForm
			case "(C)"
				PriFormIn = PriCForm
				TaskFormIn = TaskCForm
			case "(D)"
				PriFormIn = PriDForm
				TaskFormIn = TaskDForm
			case "(E)"
				PriFormIn = PriEForm
				TaskFormIn = TaskEForm
		End Select

		dueLoc = instr(1,tempStr,"due:",1)
		if dueLoc <2 then dueLoc=len(tempStr)+2
		
		'show priority
		response.write ("<td class=""td-nowrap"" width=1><center>" & PriFormIn & tempPri & "</font><a href='todotxt.asp?UpdateLine=" & server.urlencode(strLine) & "&SearchStr=" & server.urlencode(request.querystring("searchstr")) & "&PriChange=Up'>&uarr;<a href='todotxt.asp?UpdateLine=" & server.urlencode(strLine) & "&SearchStr=" & server.urlencode(request.querystring("searchstr")) & "&PriChange=Down'>&darr;</td>")
		
		' show task
		response.write ("<td width=""100%""><a href='todotxt.asp?EditLine=1&UpdateLine=" & server.urlencode(strLine) & "&SearchStr=" & server.urlencode(request.querystring("searchstr")) & "'>" & TaskFormIn & replace(replace(left(tempStr,dueLoc-2),RecStr,"")," t:"&TempThresh," ") & "</a>")
		
		'show Hide Task link
		response.write ("<a href='todotxt.asp?Hide=Hide&UpdateLine=" & server.urlencode(strLine) & "&SearchStr=" & server.urlencode(request.querystring("searchstr")) & "'>")
		response.write ("</b></font><font size=-2 style='float: right'>")
		if instr(1,strLine,"h:1",1)>1 then response.write("<b><font color=red>-</b></font>")
		response.write ( "h</a></font><br>")
		
		'show category
		response.write ("<a href=todotxt.asp?SearchStr=@" & TempCat & "><font size=-2 style='float: left'>" & TempCat & "</a>")
		
		' show recur string
		response.write ("</font><font size=-2 style='float: right'>" & right(RecStr,len(RecStr)-instr(RecStr,":")) & "</font>	</td>")
		
		' show Postpone options
		response.write ("<td nowrap=""p"" style='white-space:no-wrap;'><center>")
		response.write ("<font size=-2>")
		For Each PP in arrPP
			if PP < 8 and SearchStr=year(now())&"-"&right("0"&month(now()),2) & "-"&right("0" & day(now()),2) then 
				strPP = left(weekdayname(weekday(now()+PP)),2)
			else
				strPP = PP
			end if
			response.write ("<a href='todotxt.asp?UpdateLine=" & server.urlencode(strLine) & "&SearchStr=" & server.urlencode(request.querystring("searchstr")) & "&postpone="&PP&"&currdue=" & tempdue & "'>"&strPP&"</a> ")
		next
		response.write("</font>")
		response.write ("</td>")
		
		' show due date
		response.write ("<td nowrap=""p""")
		if cdate(tempdue)<now()-1 then response.write(" bgcolor = red")
		response.write ("><center>")
        if tempdue <> "2500-12-31" then
			response.write ("<a href='todotxt.asp?searchstr=" & tempdue & "'>")
			if cdate(tempdue)<now() then response.write("<b>")
			response.write (tempdue&"</a>")
			if TempThresh<>"2500-12-31" then response.write ("</b><br><font size=-2>t:"&TempThresh)
			response.write ("</td></tr>")
		end if
		
	else
		if (left(strline, 2)="x ") and (searchstr="Completed") then

'			dueLoc = instr(1,strline,"due:",1)
			if dueLoc > 0 then
				tempdue = mid(strline,dueLoc+4,10)
			else
				tempdue = "2500-12-31"
			end if
			if dueLoc <2 then dueLoc=len(strline)+2
		
		
			response.write ("<td></td><td><a href='todotxt.asp?EditLine=1&UpdateLine=" & server.urlencode(strLine) & "'>" & left(strLine,dueLoc-2) & "</a></td>")
			if tempdue <> "2500-12-31" then response.write ("<td nowrap=""p"">" & tempdue & "</td>")
			response.write("<tr>")
			ShowTasks = ShowTasks + 1
		end if
	end if
Next
response.write ("</table><br>")
response.write (ShowTasks & " / " &TotalTasks &" tasks<p>&nbsp;")
	
%>

</body>