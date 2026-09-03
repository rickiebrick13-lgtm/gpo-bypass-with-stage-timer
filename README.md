This is a simple bunch of files you can load into your air-gapped workstation for a countdown and flow into your presentation.

Note: 
0. Check if Python is installed, open cmd and type python --version
If it displays Python 3.x.x, proceed, else try typing py --version, if that fails too, you don't have python runtimes and will need your companies admim approval to install python or enable a local portable executable.
1. All files have to be in the same folder for simplicity.
2. In server.py, change line 7 "Test.pptx" to your "powerpoint.pptx" 
3. In launch.py, change line 7 "Test.pptx" to your "powerpoint.pptx" 
4. To alter the time, change line 40 within new Date(...) to whatever you expect, according to line 39 comment.
5. Before start, remember to close the target pptx first.
6. To start the whole thing, use CMD or PS, and cd to the folder containing the files then python server.py, and crtl+click the link. 

PS. You can use Notepad to alter the files, just no bells and whistles like IDEs. But if your company blocks Notepad then...LOL
