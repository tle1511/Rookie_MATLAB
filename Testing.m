
import mlreportgen.ppt.*
ppt = Presentation("Testing.pptx");
open(ppt);



slide1 = add(ppt,"Title and Content");
replace(slide1,"Title","Image 6");

img = Picture(which("Image6.png"));
img.Width = "5in";
img.Height = "2in";
replace(ppt,"Content",img);

close(ppt);
rptview(ppt);