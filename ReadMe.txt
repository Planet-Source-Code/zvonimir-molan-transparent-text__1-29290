AUTHOR:  Zvonimir Molan
E-mail:  zvonimir.molan@kr.hinet.hr
Web URL: http://www.inet.hr/~zmolan

The problem with standard TextBox is that it background can't be transparent, I've developed the transparent text with standard Label objects. You can't type in text in Label object at Runtime, but with help of PictureBox you can type characters and edit text in Label objects. All Label objects that you want to be transparent and editable you must put on one PictureBox, and PictureBox will then work with Label objects…
You can use the transparent text with "Microsoft Forms 2.0 Object Library" - "FM20.DLL", but there is one problem: this ActiveX component is not recommended to use, and it'll work only on computers with Microsoft Office 98 or late installed.