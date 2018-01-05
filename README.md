# todo.asp
VBScript implementation of todo.txt.

This is a very old project, but one I still use regularly.  I have been asked to share it and thought there may be more people who want it than just those few who know about it from me.  I have tried to maintain it as a desktop version of Mark Janssen's Simpletask (https://github.com/mpcjanssen/simpletask-android).  I hope to have documentation for this someday, but for now this is what you need to do:

1)  Install IIS with classic ASP.
2)  Put the files into a directory on your server.
3)  Edit the config.inc file.  You MUST supply the location for todo.txt, done.txt, and the categories you want to use.
4)  Go to http://yourserver/tododirectory/todotxt.asp

Disclaimer #1:  I cut my teeth on VBScript writing this so please keep the criticism of the coding to yourself.  If I were to rewrite it today, it would be much different (and almost definitely not an ASP using VBScript).  However, I'm not inclined to rewrite it as it works and I don't want to spend the time doing it.  If you have suggestions for bug fixes, let me know.  However, I doubt I will incorporate any suggested improvements (unless they are bug fixes).  I'll update it as I make changes.  Please feel free to make a separate branch if you'd like to continue to develop the script.

Disclaimer #2:  Backup your todo.txt and archive files file daily.  I offer this for free and as is.  It has worked great for me, but may eat your todo.txt file.
