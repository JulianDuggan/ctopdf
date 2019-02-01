SUMMARY: This do file writes a function called ctopdf that will convert a survey
formatted in a surveyCTO excel document into a much more readable PDF.  

INSTRUCTIONS: You must be using STATA15 or newer to run this program because it 
requires a command called putpdf. If you do not already have the packages egenmore
and mmerge installed, specify the packages option. Or ssc install them manually.

The basic syntax is: 

	ctopdf using "Path/to/file.xlsx", save("Path/to/save/directory") title(My Title)

Aside from the save and title options, all other options are optional. Most of the
time, you will probably want to specify the date. 

Explanations of the options: 

	* using - path to surveycto file that you would like as a PDF 
	* save - path to directory where you would like the PDFs saved
	* merge - if you would like modules to be merged into a single file. The default is 
	  to save each module of the survey as a separate pdf. 
	* skiplist - lines in the excel version of the SurveyCTO survey that you would like
	  to skip. specified as a numlist. 
	* title - title of the survey and the saved file, if you merge it automatically. 
	* date- today's date 
	* version - version number / name
	* choicelength - minimum number of choices a value label can have to be placed 
	  in the value label dictionary. All other value label options appear in the
	  text of the main survey every time they are used. 
	* coverimage - path to file of image that you would like to appear on the cover
	* translation - WORK IN PROGRESS - language that you would like translation to appear in. 
	* packages - ssc installs / updates all of the packages needed to make the program work
	* loudvars - print the name of each variable as it is formatted. mostly for debugging purposes. 
	
Fully specified, the function might look like:

	ctopdf using "Path/to/file.xlsx", save("Path/to/save/directory") merge  ///
	skiplist(100 (1) 147, 200 (1) 300) title(My Title) date(01.04.19) version(7) ///
	authors(First1 Last1, First2 Last2, and First3 Last3) choicelength(5) ///
	coverimage("Path/to/image.png") packages
 
Some things that can go wrong / that you could want to know: 

 1- If you get an error like "Failed to set table", try manually entering "putpdf clear" 3 times from the stata command line. I don't know why this works, but it sometimes does. 
 2- Make sure to use forward slashes (/) when specifying files, not backwards slashes (\)
 3- If you have questions or value labels of more than 2045 characters, they will be truncated. 
	Characters like $, }, {, and " will also be removed from all strings. 
 4- Disabled questions will not be included in the pdf.  
 5- This function will clear all of your global macros 
 6- You try to do the merge option without having installed PDFTK (right now, 
 only Windows is supported). 
 7- You don't have ssc packages mmerge and egenmore installed. Specify the packages option. 
 8- specify the DIRECTORY of the saving location, not the file name (so "C:/Users/me/Desktop" not "C:/Users/me/Desktop/survey.pdf")  
 
 To get the merge option working, install PDFTK: 
 https://www.pdflabs.com/tools/pdftk-the-pdf-toolkit/
