# Python 3
1. Another option is to use this simple batch file (save it as python34.bat or similar, but the extension must be .bat, put it wherever you like).
And then use that to run your scripts by right clicking a python script file, open with (run with) and use this batch script as default (if you want).
Also, if you have another version of Python, or is installed elsewhere, you must change the "C:\Python34\" part.
3. Install offline whl CLI: Open cmd > change to directory where whl file located:
```cmd
pip install packagename.whl -f ./ --no-index --no-deps
```
4. Chạy Pyinstaller:
```cmd
python -m PyInstaller "D:\py\main.py" --onefile
```
6. Automate SAP Python: https://k-weiming.github.io/2021-08-23-sap-connection-python/
7. Class khởi tạo với func __init__ và các func khác [VD: def buy(c,d)], khởi tạo 1 object XXX với tên_class(a,b) và gọi method lên XXX.buy
8. Tkinter: https://python-course.eu/tkinter/dialogs-in-tkinter.php
9. Saving requirement
```cmd
pip freeze my-reqs.txt
pip install -r my-reqs.txt
```
10. Cheatsheet1: https://www.dummies.com/article/technology/programming-web-design/python/python-for-kids-for-dummies-cheat-sheet-207407/
11. Cheatsheet2: https://devhints.io/python
12. SAP Scripting Tracker: https://tracker.stschnell.de/
13. SAP interesting facts: https://gayoway.com/blog/
14. Sử dụng Pyinstaller cùng với --onefile(bundles everything into a single executable) --windowed(prevents a command prompt from appearing when running GUI apps)
```cmd
pyinstaller --onefile --windowed --icon=your_icon.ico your_script.py
```
