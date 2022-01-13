import mlreportgen.ppt.*
ppt = Presentation("The Brain 3.pptx");
open(ppt);

cd = niftiread("cd.nii.gz");
ki = niftiread("ki67.nii.gz");
tc = niftiread("tc.nii.gz");


fig = figure;

%% 
% Adding the slides

slide1 = add(ppt,"Title Slide");
replace(slide1,"Title","The Brain 3");

slide2 = add(ppt,"Two Content");
replace(slide2,"Title","Image 40");

cd40 = imagesc(cd(:,:,40));
colormap("gray");
snapshotcd40 = "cd40.png";
print(fig,"-dpng","cd40.png");
figcd40 = Picture("cd40.png");

ki40 = imagesc(ki(:,:,40));
snapshotki40 = "ki40.png";
print(fig,"-dpng","ki40.png");
figki40 = Picture("ki40.png");

replace(slide2,"Left Content", figcd40);

replace(slide2,"Right Content",figki40);

%% 
% Close and review

close(ppt);
rptview(ppt);