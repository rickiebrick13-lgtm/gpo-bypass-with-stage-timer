This is a simple bunch of files you can load into your air-gapped workstation for a countdown and flow into your presentation.

Critical:
1. Check if Python is installed, open cmd and type python --version
2. If it displays Python 3.x.x, proceed, else try typing py --version, if that fails too, you don't have python runtimes and will need your company admin approval to install python or enable a local portable executable.
3. If you're out of luck, then use the files in the VBSclock folder.

Note (python version):
0. Use the files in pyClocker folder.
1. All files have to be in the same folder for simplicity.
2. In server.py, change line 7 "Test.pptx" to your "powerpoint.pptx" 
3. In launch.py, change line 7 "Test.pptx" to your "powerpoint.pptx" 
4. To alter the time, change line 40 within new Date(...) to whatever you expect, according to line 39 comment.
5. Before start, remember to close the target pptx first.
6. To start the whole thing, use CMD or PS, and cd to the folder containing the files then python server.py, and crtl+click the link.
7. For emergency launch, hit the spacebar.

Note (vbscript version):
0. Use the files in the VBSclock folder.
1. All files have to be in the same folder
2. In timeScript.vbs, change TargetTime = "YYYY-MM-DD HH:MM:SS" to alter the time
3. In timeScript.vbs, change PresentationFile = "powerpoint.pptx" your powerpoint presentation's name.
4. After double clicking timeScript.vbs, it will run in the background, just open index.html, when the timer hits 0, the powerpoint will replace the browser and be on top of it in full-screen mode.

PS. You can use Notepad to alter the files, just no bells and whistles like IDEs. But if your company blocks Notepad then...LOL
