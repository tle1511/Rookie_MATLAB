nii = niftiread("ICBM_Template\ICBM_Template.nii");
import mlreportgen.ppt.*;
ppt = Presentation("ICBM_Practice1.pptx");
open(ppt);

slide1 = add(ppt,"Title Slide");
replace(slide1,"Title","MRI of the Brain");

slide2 = add(ppt,"Title and Picture");
img1to5 = imagesc(nii(:,:,1));
gscale1to5 = imadjust(img1to5);
replace(slide2,"Title","Image 1 to 5");
replace(slide2,"Picture",gscale1to5);

slide3 = add(ppt,"Title and Picture");
img6 = imagesc(nii(:,:,6));
gscale6 = imadjust(img6);
replace(slide3,"Title","Image 6");
replace(slide3,"Picture",gscale6);

slide4 = add(ppt,"Title and Picture");
img7 = imagesc(nii(:,:,7));
gscale7 = imadjust(img7);
replace(slide4,"Title","Image 7");
replace(slide4,"Picture",gscale7);

slide5 = add(ppt,"Title and Picture");
img8 = imagesc(nii(:,:,8));
gscale8 = imadjust(img8);
replace(slide5,"Title","Image 8");
replace(slide5,"Picture",gscale8);

slide6 = add(ppt, "Title and Picture");
img9 = imagesc(nii(:,:,9));
gscale9 = imadjust(img9);
replace(slide6,"Title","Image 9");
replace(slide6,"Picture",gscale9);

slide7 = add(ppt, "Title and Picture");
img10 = imagesc(nii(:,:,10));
gscale10 = imadjust(img10);
replace(slide7,"Title","Image 10");
replace(slide7,"Picture",gscale10);


slide8 = add(ppt, "Title and Picture");
img11 = imagesc(nii(:,:,11));
gscale11 = imadjust(img11);
replace(slide8,"Title","Image 11");
replace(slide8,"Picture",gscale11);


slide9 = add(ppt, "Title and Picture");
img12 = imagesc(nii(:,:,12));
gscale12 = imadjust(img12);
replace(slide9,"Title","Image 12");
replace(slide9,"Picture",gscale12);

slide10 = add(ppt, "Title and Picture");
img13 = imagesc(nii(:,:,13));
gscale13 = imadjust(img13);
replace(slide10,"Title","Image 13");
replace(slide10,"Picture",gscale13);


slide11 = add(ppt, "Title and Picture");
img14 = imagesc(nii(:,:,14));
gscale14 = imadjust(img14);
replace(slide11,"Title","Image 14");
replace(slide11,"Picture",gscale14);


slide12 = add(ppt, "Title and Picture");
img15 = imagesc(nii(:,:,15));
gscale15 = imadjust(img15);
replace(slide12,"Title","Image 15");
replace(slide12,"Picture",gscale15);


slide13 = add(ppt, "Title and Picture");
img16 = imagesc(nii(:,:,16));
gscale16 = imadjust(img16);
replace(slide13,"Title","Image 16");
replace(slide13,"Picture",gscale16);


slide14 = add(ppt, "Title and Picture");
img17 = imagesc(nii(:,:,17));
gscale17 = imadjust(img17);
replace(slide14,"Title","Image 17");
replace(slide14,"Picture",gscale17);


slide15 = add(ppt, "Title and Picture");
img18 = imagesc(nii(:,:,18));
gscale18 = imadjust(img18);
replace(slide13,"Title","Image 18");
replace(slide13,"Picture",gscale18);


slide16 = add(ppt, "Title and Picture");
img19 = imagesc(nii(:,:,19));
gscale19 = imadjust(img19);
replace(slide16,"Title","Image 19");
replace(slide16,"Picture",gscale19);

slide17 = add(ppt, "Title and Picture");
img20 = imagesc(nii(:,:,20));
gscale20 = imadjust(img20);
replace(slide17,"Title","Image 20");
replace(slide17,"Picture",gscale20);

slide18 = add(ppt, "Title and Picture");
img21 = imagesc(nii(:,:,21));
gscale21 = imadjust(img21);
replace(slide18,"Title","Image 21");
replace(slide18,"Picture",gscale21);

slide19 = add(ppt, "Title and Picture");
img22 = imagesc(nii(:,:,22));
gscale22 = imadjust(img22);
replace(slide13,"Title","Image 22");
replace(slide13,"Picture",gscale22);

slide20 = add(ppt, "Title and Picture");
img23 = imagesc(nii(:,:,23));
gscale23 = imadjust(img23);
replace(slide20,"Title","Image 23");
replace(slide20,"Picture",gscale23);

slide21 = add(ppt, "Title and Picture");
img24 = imagesc(nii(:,:,24));
gscale24 = imadjust(img24);
replace(slide21,"Title","Image 24");
replace(slide21,"Picture",gscale24);

slide22 = add(ppt, "Title and Picture");
img25 = imagesc(nii(:,:,25));
gscale25 = imadjust(img25);
replace(slide22,"Title","Image 25");
replace(slide22,"Picture",gscale25);

slide23 = add(ppt, "Title and Picture");
img26 = imagesc(nii(:,:,26));
gscale26 = imadjust(img26);
replace(slide23,"Title","Image 26");
replace(slide23,"Picture",gscale26);

slide24 = add(ppt, "Title and Picture");
img27 = imagesc(nii(:,:,27));
gscale27 = imadjust(img27);
replace(slide24,"Title","Image 27");
replace(slide24,"Picture",gscale27);

slide25 = add(ppt, "Title and Picture");
img28 = imagesc(nii(:,:,28));
gscale28 = imadjust(img28);
replace(slide25,"Title","Image 28");
replace(slide25,"Picture",gscale28);

slide26 = add(ppt, "Title and Picture");
img29 = imagesc(nii(:,:,29));
gscale29 = imadjust(img29);
replace(slide26,"Title","Image 29");
replace(slide26,"Picture",gscale29);

slide27 = add(ppt, "Title and Picture");
img30 = imagesc(nii(:,:,30));
gscale30 = imadjust(img30);
replace(slide27,"Title","Image 30");
replace(slide27,"Picture",gscale30);

slide28 = add(ppt, "Title and Picture");
img31 = imagesc(nii(:,:,31));
gscale31 = imadjust(img31);
replace(slide28,"Title","Image 31");
replace(slide28,"Picture",gscale31);

slide29 = add(ppt, "Title and Picture");
img32 = imagesc(nii(:,:,32));
gscale32 = imadjust(img32);
replace(slide29,"Title","Image 32");
replace(slide29,"Picture",gscale32);

slide30 = add(ppt, "Title and Picture");
img33 = imagesc(nii(:,:,33));
gscale33 = imadjust(img33);
replace(slide30,"Title","Image 33");
replace(slide30,"Picture",gscale33);

slide31 = add(ppt, "Title and Picture");
img34 = imagesc(nii(:,:,34));
gscale34 = imadjust(img34);
replace(slide31,"Title","Image 34");
replace(slide31,"Picture",gscale34);

slide32 = add(ppt, "Title and Picture");
img35 = imagesc(nii(:,:,35));
gscale35 = imadjust(img35);
replace(slide32,"Title","Image 35");
replace(slide32,"Picture",gscale35);

slide33 = add(ppt, "Title and Picture");
img36 = imagesc(nii(:,:,36));
gscale36 = imadjust(img36);
replace(slide33,"Title","Image 36");
replace(slide33,"Picture",gscale36);

slide34 = add(ppt, "Title and Picture");
img37 = imagesc(nii(:,:,24));
gscale24 = imadjust(img24);
replace(slide21,"Title","Image 24");
replace(slide21,"Picture",gscale24);

slide35 = add(ppt, "Image 38");
img24 = imagesc(nii(:,:,24));
gscale24 = imadjust(img24);
replace(slide21,"Title","Image 24");
replace(slide21,"Picture",gscale24);

slide36 = add(ppt, "Image 39");
img24 = imagesc(nii(:,:,24));
gscale24 = imadjust(img24);
replace(slide21,"Title","Image 24");
replace(slide21,"Picture",gscale24);

slide37 = add (ppt, "Image 40");
img24 = imagesc(nii(:,:,24));
gscale24 = imadjust(img24);
replace(slide21,"Title","Image 24");
replace(slide21,"Picture",gscale24);

slide38 = add(ppt, "Image 41");
img24 = imagesc(nii(:,:,24));
gscale24 = imadjust(img24);
replace(slide21,"Title","Image 24");
replace(slide21,"Picture",gscale24);

slide39 = add(ppt, "Image 42");
img24 = imagesc(nii(:,:,24));
gscale24 = imadjust(img24);
replace(slide21,"Title","Image 24");
replace(slide21,"Picture",gscale24);

slide40 = add(ppt, "Image 43");
img24 = imagesc(nii(:,:,24));
gscale24 = imadjust(img24);
replace(slide21,"Title","Image 24");
replace(slide21,"Picture",gscale24);

slide41 = add(ppt, "Image 44");
img24 = imagesc(nii(:,:,24));
gscale24 = imadjust(img24);
replace(slide21,"Title","Image 24");
replace(slide21,"Content",gscale24);

slide42 = add(ppt, "Image 45");
img24 = imagesc(nii(:,:,24));
gscale24 = imadjust(img24);
replace(slide21,"Title","Image 24");
replace(slide21,"Content",gscale24);

slide43 = add(ppt, "Image 46");
img24 = imagesc(nii(:,:,24));
gscale24 = imadjust(img24);
replace(slide21,"Title","Image 24");
replace(slide21,"Content",gscale24);

slide44 = add(ppt, "Image 47");
img24 = imagesc(nii(:,:,24));
gscale24 = imadjust(img24);
replace(slide21,"Title","Image 24");
replace(slide21,"Content",gscale24);

slide45 = add(ppt, "Image 48");
img24 = imagesc(nii(:,:,24));
gscale24 = imadjust(img24);
replace(slide21,"Title","Image 24");
replace(slide21,"Content",gscale24);

slide46 = add(ppt, "Image 49");
img24 = imagesc(nii(:,:,24));
gscale24 = imadjust(img24);
replace(slide21,"Title","Image 24");
replace(slide21,"Content",gscale24);

slide47 = add(ppt, "Image 50");
img24 = imagesc(nii(:,:,24));
gscale24 = imadjust(img24);
replace(slide21,"Title","Image 24");
replace(slide21,"Content",gscale24);

slide48 = add(ppt, "Image 51");
img24 = imagesc(nii(:,:,24));
gscale24 = imadjust(img24);
replace(slide21,"Title","Image 24");
replace(slide21,"Content",gscale24);

slide49 = add(ppt, "Image 52");
img24 = imagesc(nii(:,:,24));
gscale24 = imadjust(img24);
replace(slide21,"Title","Image 24");
replace(slide21,"Content",gscale24);

slide50 = add(ppt, "Image 53");
img24 = imagesc(nii(:,:,24));
gscale24 = imadjust(img24);
replace(slide21,"Title","Image 24");
replace(slide21,"Content",gscale24);

slide51 = add(ppt, "Image 54");

slide52 = add(ppt, "Image 55");

slide53 = add(ppt, "Image 56");

slide54 = add(ppt, "Image 57");

slide55 = add(ppt, "Image 58");

slide56 = add(ppt, "Image 59");

slide57 = add(ppt, "Image 60");

slide58 = add(ppt, "Image 61");

slide59 = add(ppt, "Image 62");

slide60 = add(ppt, "Image 63");

slide61 = add(ppt, "Image 64");

slide62 = add(ppt, "Image 65");

slide63 = add(ppt, "Image 66");

slide64 = add(ppt, "Image 67");

slide65 = add(ppt, "Image 68");

slide66 = add(ppt, "Image 69");

slide67 = add(ppt, "Image 70");

slide68 = add(ppt, "Image 71");

slide69 = add(ppt, "Image 72");

slide70 = add(ppt, "Image 73");

slide71 = add(ppt, "Image 74");

slide72 = add(ppt, "Image 75");

slide73 = add(ppt, "Image 76");

slide74 = add(ppt, "Image 77");

slide75 = add(ppt, "Image 78");

slide76 = add(ppt, "Image 79");

slide77 = add(ppt, "Image 80");

slide78 = add(ppt, "Image 81");

slide79 = add(ppt, "Image 82");

slide80 = add(ppt, "Image 83");

slide81 = add(ppt, "Image 84");

slide82 = add(ppt, "Image 85");

slide83 = add(ppt, "Image 86");

slide84 = add(ppt, "Image 87");

slide85 = add(ppt, "Image 88");

slide86 = add(ppt, "Image 89");

slide87 = add(ppt, "Image 90");

slide88 = add(ppt, "Image 91");

slide89 = add(ppt, "Image 92");

slide90 = add(ppt, "Image 93");

slide91 = add(ppt, "Image 94");

slide92 = add(ppt, "Image 95");

slide93 = add(ppt, "Image 96");

slide94 = add(ppt, "Image 97");

slide95 = add(ppt, "Image 98");

slide96 = add(ppt, "Image 99");

slide97 = add(ppt, "Image 100");

slide98 = add(ppt, "Image 101");

slide99 = add(ppt, "Image 102");

slide100 = add(ppt, "Image 103");

slide101 = add(ppt, "Image 104");

slide102 = add(ppt, "Image 105");

slide103 = add(ppt, "Image 106");

slide104 = add(ppt, "Image 107");

slide105 = add(ppt, "Image 108");

slide106 = add(ppt, "Image 109");

slide107 = add(ppt, "Image 110");

slide108 = add(ppt, "Image 111");

slide109 = add(ppt, "Image 112");

slide110 = add(ppt, "Image 113");

slide111 = add(ppt, "Image 114");

slide112 = add(ppt, "Image 115");

slide113 = add(ppt, "Image 116");

slide114 = add(ppt, "Image 117");

slide115 = add(ppt, "Image 118");

slide116 = add(ppt, "Image 119");

slide117 = add(ppt, "Image 120");

slide118 = add(ppt, "Image 121");

slide119 = add(ppt, "Image 122");

slide120 = add(ppt, "Image 123");

slide121 = add(ppt, "Image 124");

slide122 = add(ppt, "Image 125");

slide123 = add(ppt, "Image 126");

slide124 = add(ppt, "Image 127");

slide125 = add(ppt, "Image 128");

slide126 = add(ppt, "Image 129");

slide127 = add(ppt, "Image 130");

slide128 = add(ppt, "Image 131");

slide129 = add(ppt, "Image 132");

slide130 = add(ppt, "Image 133");

slide131 = add(ppt, "Image 134");

slide132 = add(ppt, "Image 135");

slide133 = add(ppt, "Image 136");

slide134 = add(ppt, "Image 137");

slide135 = add(ppt, "Image 138");

slide136 = add(ppt, "Image 139");

slide137 = add(ppt, "Image 140");

slide138 = add(ppt, "Image 141");

slide139 = add(ppt, "Image 142");

slide140 = add(ppt, "Image 143");

slide141 = add(ppt, "Image 144");

slide142 = add(ppt, "Image 145");

slide143 = add(ppt, "Image 146");

slide144 = add(ppt, "Image 147");

slide145 = add(ppt, "Image 148");

slide146 = add(ppt, "Image 149");

slide147 = add(ppt, "Image 150");

slide148 = add(ppt, "Image 151");

slide149 = add(ppt, "Image 152");

slide150 = add(ppt, "Image 153");

slide151 = add(ppt, "Image 154");

slide152 = add(ppt, "Image 155");

slide153 = add(ppt, "Image 156");

slide154 = add(ppt, "Image 157");

slide155 = add(ppt, "Image 158");

slide156 = add(ppt, "Image 159");

slide157 = add(ppt, "Image 160-181")

close(ppt);
rptview(ppt);
%% 
% 



%% 
%