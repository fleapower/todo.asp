<%
' todo file
ToDoFile = "todo.txt"

' archive file
ArchFile = "done.txt"

' Put categories here
ArrCats=Array("Anywhere","Home","Work","Shopping")

' ArrDueD is the list of dates you want to use for initial date.  The entries are in pairs.  First put the text of the selection, then the number of days until due.  Note:  "None" uses "None" for number of days until due so if you use "None," it should be followed by "None."
ArrDueD=Array("None","None","Today",0,"Tomorrow",1,"One Week",7,"Two Weeks",14,"30 Days",30)

' This is the day which will be selected by default for the due date on new tasks.
SelectedDue = "Today"

' put days to postpone here
ArrPP=Array(1,2,3,4,5,6,7,14,30)

' priority formats
PriAForm = "<b><font color=red>"
PriBForm = "<font color=orange>"
PriCForm = "<font color=brown>"
PriDForm = "<font color=blue>"
PriEForm = "<font color=purple>"

' task text formats based on priority
TaskAForm = "<b>"
TaskBForm = "<b>"
TaskCForm = ""
TaskDForm = ""
TaskEForm = ""

' Theme. Uncomment theme to be used (DARK or LIGHT - use caps)
todotxtTheme = "LIGHT"
%>
