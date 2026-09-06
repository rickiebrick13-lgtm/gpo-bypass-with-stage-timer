// launch.js - Air-Gapped JScript Timer & PPT Launcher
var targetTime = new Date("2026-09-06T16:00:00").getTime(); // Target T-0

while (new Date().getTime() < targetTime) {
    WScript.Sleep(1000); // Poll every second without burning CPU
}

// T-0 Hit: Fire PowerPoint
var ppt = new ActiveXObject("PowerPoint.Application");
ppt.Visible = true;

// Open presentation and trigger full-screen slideshow
var presentation = ppt.Presentations.Open("C:\\path\\to\\presentation.pptx");
presentation.SlideShowSettings.Run();