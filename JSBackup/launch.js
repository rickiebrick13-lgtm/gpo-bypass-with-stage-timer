// launch.js
var ppt = new ActiveXObject("PowerPoint.Application");
ppt.Visible = true;
ppt.Presentations.Open("C:\\path\\to\\presentation.pptx");
ppt.ActivePresentation.SlideShowSettings.Run();