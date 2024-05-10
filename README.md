# Python
Another option is to use this simple batch file (save it as python34.bat or similar, but the extension must be .bat, put it wherever you like):

@ECHO OFF
C:\Python34\python.exe %*
pause
@ECHO ON

And then use that to run your scripts by right clicking a python script file, open with (run with) and use this batch script as default (if you want). Also, if you have another version of Python, or is installed elsewhere, you must change the "C:\Python34\" part.
