/*
SUMMARY: This do file writes a function called ctopdf that will convert a survey
formatted in a surveyCTO excel document into a much more readable PDF.  

INSTRUCTIONS: You must be using STATA15 or newer to run this program because it 
requires a command called putpdf. If you do not already have the packages egenmore
and mmerge installed, specify the packages option. Or ssc install them manually.

The basic syntax is: 

	ctopdf using "Path/to/file.xlsx", save("Path/to/save/directory"). 

Aside from the save option, all other options are optional. Most of the
time, you will probably want to specify the title of your survey and the date. 

Explanations of the options: 

	* using - path to surveycto file that you would like as a PDF 
	* save - path to directory where you would like the PDFs saved
	* merge - if you would like modules to be merged into a single file. The default is 
	  to save each module of the survey as a separate pdf. 
	* skiplist - lines in the excel version of the SurveyCTO survey that you would like
	  to skip. specified as a numlist. 
	* title - title of the survey
	* date- today's date 
	* version - version number / name
	* choicelength - minimum number of choices a value label can have to be placed 
	  in the value label dictionary. All other value label options appear in the
	  text of the main survey every time they are used. 
	* coverimage - path to file of image that you would like to appear on the cover
	* translation - language that you would like translation to appear in. WORK IN PROGRESS

Fully specified, the function might look like:

	ctopdf using "Path/to/file.xlsx", save("Path/to/save/directory") merge  ///
	skiplist(100 (1) 147, 200 (1) 300) title(My Title) date(01.04.19) version(7) ///
	authors(First1 Last1, First2 Last2, and First3 Last3) choicelength(5) ///
	translation(swahili) coverimage("Path/to/image.png")
 
Some things that can go wrong / that you could want to know: 
 1- If you get an error like "Failed to set table", try manually entering "putpdf clear" 3 times from the stata command line. I don't know why this works, but it sometimes does. 
 2- Make sure to use forward slashes (/) when specifying files, not backwards slashes (\)
 3- If you have questions or value labels of more than 2045 characters, they will be truncated. 
	Characters like $, }, {, and " will also be removed from all strings. 
 4- Disabled questions will not be included in the pdf.  
 5- This function will clear all of your global macros 
 
 */

cap prog drop ctopdf 
prog def ctopdf 
syntax using/, SAve(string)[Merge SKiplist(numlist) TItle(string) Date(string) Packages TRANSLation Version(string) CHOICElength(integer 10) COVERimage(string) Authors(string)]

////////////////////////////////////////////////////////////////////////////////
/////////////////////////////// Preliminaries //////////////////////////////////
////////////////////////////////////////////////////////////////////////////////

// clean slate
clear 
clear mata
putpdf clear 
putpdf clear 
putpdf clear
* macro drop doc_count table_count question_count heading* rheading*  
set more off 
set maxvar 30000

// make sure any idiosyncratic packages that we use are installed 
capture confirm e `packages' 
if !_rc {
	foreach package in mmerge egenmore {
		capture which `package' 
		if _rc == 111 ssc install `package'
	}
}
	
////////////////////////////////////////////////////////////////////////////////
/////////////////////////////// Clean Survey ///////////////////////////////////
////////////////////////////////////////////////////////////////////////////////
					
// CHOICES SHEET

	// import choice list 
	qui import excel using "`using'", clear firstrow sheet("choices")
	
	// standardize variable names / set
	* NOTE: IF YOU WANT TO ADD TRANSLATIONS, START BY CHANGING NEXT LINE 
	capture confirm var name , exact
	if !_rc rename name value
	capture keep list_name value label
	if _rc == 111 {
		di "Make sure that your choice sheet's column names are list_name, value, and label." 
		exit
	}
	
	// gen / clean vars long
	qui replace list_name = strtrim(list_name)
	qui replace label = subinstr(label, `"""', "", .)  
	qui sort list_name value 
	qui bysort list_name: gen choice_count = _n  
	qui drop if list_name == ""
	qui drop if list_name == " "  
	
	// reshape wide
	qui reshape wide label value, i(list_name) j(choice_count)
	
	// gen / clean vars wide
	foreach l of varlist label* {
		qui replace `l' = `"""' + `l' + `"" "' if !mi(`l')
	}	
	qui egen choices = concat(label*)
	
	foreach l of varlist value* {
		qui tostring `l', replace
		qui replace `l' = "" if `l' == "."
		qui replace `l' = `"""' + `l' + `"" "' if !mi(`l')
	}
	qui egen values = concat(value*)
	qui egen choicecount = nwords(choices)
	qui egen valuecount = nwords(values)	
	
	// check for errors
	capture assert choicecount == valuecount
	if _rc == 9 {
		ta list_name if choicecount != valuecount, m
		di "Warning: Some of your value labels may not correspond to the appropriate values. Check for weird characters in your value labels" 
	}
	
	// standardize var set 
	qui keep list_name choices choicecount values
	
	// save 
	tempfile choices
	qui save `choices' 

// SURVEY SHEET

	// import survey 
	qui import excel using "`using'", clear firstrow sheet("survey")
	
	// gen index 
	* NOTE: We create the below index immediately so that no subsequent commands alter 
	* the order in which questions appear
	qui gen index = _n
	
	// standardize variable set 
	loc actualvars
	local possiblevars index type name label* hint* default appearance constraint* relevance disabled required* readonly calculation repeat_count media* choice_filter note response_note publishable minimum_seconds
	foreach a in `possiblevars' {
		capture ds `a'
		if _rc == 0 loc actualvars = `"`actualvars' `a'"'
	}
	keep `actualvars'

	// gen / clean other key vars 
	
		// names 
		qui duplicates tag name if name != "" & strpos(type, "begin ") == 0 & strpos(type, "end ") == 0 , gen(dupname)		
		capture assert dupname == 0 | dupname == .
		if _rc == 9 {
			ta name dupname if dupname !=. & dupname !=0, m
			di "Some of your questions(other than begin group, end group, begin repeat, end repeat) have duplicated names -- remove the duplicates!" 
			drop dupname
			exit
		}
		drop dupname 
		
		// group 
		qui egen group_num = group2(name) if strpos(type, "group") > 0, sort(min(index)) label
		qui gen group = ""
		qui sum group_num
		forvalues i = 1 / `r(max)' { 
			qui levelsof name if group_num == `i', loc(groupname) clean 
			qui sum index if group_num == `i', d 
			qui replace group = "`groupname'" if index <=`r(max)' & index >= `r(min)'
		}
	
		// choice label name
		qui split type if strpos(type, "select_one") + strpos(type, "select_multiple") > 0, gen(choices)
		qui drop choices1
		qui rename choices2 list_name
	
		// type (excludes value label name) 
		qui gen type_clean = word(type, 1) 		
		qui replace type = type_clean if strpos(type, "select_one") + strpos(type, "select_multiple") > 0

// MERGED SURVEY AND CHOICES

	// merge
	qui mmerge list_name using `choices', unmatched(master)
	qui drop _m 
	
	// sort by index
	qui sort index

	// clean string variables 
	qui ds, has(type str# strL)
	foreach var in `r(varlist)' {
		
		// no leading and lagging spaces
		qui replace `var' = strtrim(`var') 
		
		// replace strL (long strings) with str1-2045
		capture confirm strL variable `var'
		if !_rc {
			tempvar l`var'
			qui gen `l`var'' = length(`var') 
			qui sum `l`var''
			loc len = `r(max)'
			
			// throw warning for truncated vars 
			if `len' >=2045 {
				loc truncated = `len' - 2045
				di "Warning: Some values of `var' have been truncated by up to `truncated' characters."
			}
			
			loc len = min(`len', 2045)
			loc len = max(1, `len')
			qui recast str`len' `var', force 
		}
		
		// get rid of $, }, {, and " which STATA will misinterpret as referring to globals, defining loops, delimiting strings 
		capture confirm str# variable `var' 
			if !_rc {
				qui replace `var' = subinstr(`var', "$", "", .)
				qui replace `var' = subinstr(`var', "{", "", .)
				qui replace `var' = subinstr(`var', "}", "", .) 
				if "`var'" != "choices" & "`var'" != "values" {
					qui replace `var' = subinstr(`var', `"""', "", .) 
					
				}
			}	
	}
	
////////////////////////////////////////////////////////////////////////////////
///////////////////////////////// format survey ////////////////////////////////
////////////////////////////////////////////////////////////////////////////////							
				
	// count "real" questions during loop
	* NOTE: types included are text, integer, select_one, and select_mulitple
	global question_count = 0
	* NOTE: if in later updates, a single loop of the formatquestion function can 
	* create more than 10 tables, then you should  increase table_count
	global table_count = 10 
	global doc_count = 0

	// begin 
	putpdf clear 
	putpdf begin, font(, 10) margin(left, .1in) margin(right, .1in) 
	
	// title page
	putpdf paragraph, font(, 16) halign(center) 
	putpdf text ("`title'"), linebreak(3) bold 
	putpdf text ("`date'"), linebreak(1)
	
	capture confirm e `version'
	if !_rc putpdf text ("Version: `version'"), linebreak(3)
	
	capture confirm e `coverimage' 
	if !_rc {
		putpdf paragraph, halign(center)
		putpdf image `coverimage', linebreak(3)
	}
	
	capture confirm e `authors' 
	if !_rc {
		putpdf paragraph, halign(center)
		putpdf text ("By: `authors'") 
	}
	
	putpdf pagebreak 

	// main survey 
	qui sum index 
	forvalues i = 1/`r(max)' {

	// ignore disabled questions

		qui levelsof disable if index == `i', loc(disabled) clean
		qui levelsof type if index == `i', loc(type) clean 
		
		if "`disabled'" == "yes" & "`type'" != "MODULE" continue 
		
		// ignore questions on the skip list (because they are repeats, or for some other reason) 
		loc skip = 0
		capture confirm e `skiplist' 
		if !_rc {
			foreach j of numlist `skiplist' {
				loc k = `i' - 1
				if `k' == `j' loc skip = `skip' + 1
				else loc skip = `skip' + 0
		}
		if `skip' == 1 {
			di "skipped row `i'"
			continue
		}
		}
		
		// save modules and large sections (>=500 tables) individually 
		if $table_count >=500 | ("`type'" == "MODULE" & $doc_count > 0) {
			global doc_count = $doc_count + 1
			if ${doc_count} < 10 loc docstr = "0${doc_count}"
			else loc docstr = "${doc_count}"

			if "`type'" == "MODULE" { 
				qui egen modulename = sieve(name), keep(alphabetic numeric) 
				qui levelsof modulename if index == `i', loc(modulename) clean
				qui drop modulename
				di "`save'/part`docstr'_`modulename'.pdf"
				putpdf save "`save'/part`docstr'_`modulename'.pdf", replace
			}
			else putpdf save "`save'/part`docstr'_`modulename'.pdf", replace
			di "saved part ${doc_count}, `modulename'"

			putpdf clear 
			putpdf begin 
			global table_count = 10 
		}
	
		preserve 
		qui keep if index == `i'
		formatquestion
		restore 
	}
	
	di "just before final"
	
	// save final survey section
	global doc_count = $doc_count + 1
	if ${doc_count} < 10 loc docstr = "0${doc_count}"
	else loc docstr = "${doc_count}"	
	putpdf save "`save'/part`docstr'_`modulename'.pdf", replace
	di "saved part ${doc_count}"
	
	// value label index 
	putpdf clear 
	putpdf begin , font(, 10) margin(left, .1in) margin(right, .1in) 
	putpdf paragraph, font(, 14) halign(center) 
	putpdf text ("Value Label Dictionary"), bold
	levelsof list_name if choicecount >= `choicelength', loc(longvls) clean 
	capture confirm e `longvls' 
	if !_rc {
	foreach l in `longvls' {
		qui levelsof choices if list_name == "`l'", loc(choices) clean
		qui levelsof values if list_name == "`l'", loc(values) clean 
		qui levelsof choicecount if list_name == "`l'", loc(ccount) clean
		putpdf paragraph, halign(center) 
		putpdf text ("`l'"), bold
		putpdf table vl = (`ccount', 2), halign(center)
		forvalues j = 1 / `ccount' { 
			loc w : word `j' of `choices'
			loc v: word `j' of `values'
			putpdf table vl(`j', 1) = ("`w'")
			putpdf table vl(`j', 2) = ("`v'") 
			
		}
	}
	global doc_count = $doc_count + 1
	if ${doc_count} < 10 loc docstr = "0${doc_count}"
	else loc docstr = "${doc_count}"	
	putpdf save "`save'/part`docstr'_valuelabels.pdf", replace 
	di "saved part ${doc_count}, value label dictionary"
	}
	
// merge files 
	capture confirm e `merge'
	if !_rc {
		shell
		shell "S:\Dropbox\SurveyCTO_to_PDF\Code\merge.py"
	} 
	else{
		di "merge the files manually!" 
	}
	
end

////////////////////////////////////////////////////////////////////////////////
///////////////////// build question formating function ////////////////////////
////////////////////////////////////////////////////////////////////////////////

cap prog drop formatquestion
prog def formatquestion

	// define fundamental locals 
	* NOTE: currently excluded: default  
	* requiredmessage readonly mediaimage mediaaudio mediavideo  note
	* choice_filter response_note publishable labelswahili constraintmessageswahili		
	* NOTE: disabled field is handled outside this function 
	local ctovars index type name label hint constraint relevance required ///
	choices values choicecount group list_name calculation constraintmessage ///
	repeat_count
	foreach var in `ctovars' {
		 qui levelsof `var', loc(`var') clean
	}

	// clean up a few locals 
		
		// group, relevance, constraint, and required locals 
		
			// number of rows in table containing this info 
			loc rows = 1
			
			// group 
			capture confirm e `group' 
			if !_rc { 
				loc group_details = "`group'"
				loc rows = `rows' + 1 
			}	
			else {
				loc group_details 
			}
			
			// name 
			capture confirm e `name' 
			if !_rc { 
				loc name_details = "`name'"
				loc rows = `rows' + 1 
			}	
			else {
				loc name_details 
			}
			
			// relevance
			capture confirm e `relevance'
			if !_rc {
				loc relevance_details = `"`relevance'"'
				loc rows = `rows' + 1
			}
			else {
				loc relevance_details
			}
	
			// constraint
			capture confirm e `constraint'
			if !_rc {
				loc constraint_details = `"`constraint'"'
				loc rows = `rows' + 1					
			}
			else {
				loc constraint_details
			}
						
			// required 
			* NOTE: "YES" is the default rather than "NO" 
			capture confirm e `required' 
			if !_rc {
				loc required_details = upper(`"`required'"')
				if "`required_details'" == "YES" {
					loc requied_details 
				}
			}
			else {
				loc required_details = "NO" 
				loc rows = `rows' + 1
			}
		
		// question label and hint 
			
			// hint
			capture confirm e `hint' 
			if !_rc loc hint_details = (`"  (Hint: `hint')"')
			else loc hint_details
			
			// labels 
			capture confirm e `label' 
			if !_rc loc label_details = "`label'"
			else {
				if "`type'" != "calculate" & "`type'" != "calculate_here"  ///
				loc label_details = "THIS FIELD HAS NO TEXT"
				else loc label_details
			}

		// constraint message 
		capture confirm e `constraintmessage' 
		if !_rc loc constraintmessage_details = `"`constraintmessage'"'
		else loc constraintmessage_details 

		// calculation 
		capture confirm e `calculation' 
		if !_rc loc calculation_details = "`calculation'" 
		else loc calculation_details = "NO CALCULATION IS PERFORMED HERE" 
		
		
		// repeat count 
		capture confirm e `repeat_count' 
		if !_rc loc repeatcount_details = "`repeatcount'"
		else loc repeatcount_details = "NO REPEAT COUNT GIVEN"

	// define / perform formatting for each type of question 
	if "`type'" != "" | "`name'" != "" di "`type', `name'" 

		// BEGIN MODULE
		if "`type'" == "MODULE" {
			putpdf paragraph, halign(center) font(,16) 
			putpdf text (`"Module: `name'"'), bold linebreak(3) 
			
		}
		
		// BEGIN GROUP
		if "`type'" == "begin group" {
			
			// define global (to be used with footer, local won't survive) 
			capture confirm e `label' 
			if !_rc {
				global heading`name' = "`label_details'" 
			}
			else {
				global heading`name' = "`name'"
			}
			
			// put in pdf 
			putpdf paragraph, font(,12)
			putpdf text (" "), linebreak(2)
			putpdf text (`"Begin Group: ${heading`name'}"'), bold 
			putpdf text ( "- `name'") 
			
			capture confirm e `relevance_details' 
			if !_rc {
				loc relevance_details = "`relevance_details'" 
			}
			else {
				loc relevance_details = "all"
			}
			putpdf paragraph
			putpdf text (`"Section relevance: `relevance_details'"'), linebreak(2) italic
		}

		// END GROUP
		if "`type'" == "end group" {

			// define locals
			loc footing = `"${heading`name'}"'
				
			// put in pdf 
			putpdf paragraph, font(,12) 
			putpdf text ("End Group: `footing'"), bold linebreak(2)
		}
		
		
		// BEGIN REPEAT
		if "`type'" == "begin repeat" {
			
			// define global (to be used with footer, local won't survive) 
			capture confirm e `label' 
			if !_rc {
				global rheading`name' = "`label_details'" 
			}
			else {
				global rheading`name' = "`name'"
			}
			
			// put in pdf 
			putpdf paragraph, font(,12)
			putpdf text (" "), linebreak(2)
			putpdf text (`"Begin Repeat: ${rheading`name'}"'), bold 
			putpdf text ( "- `name'") 
			
			capture confirm e `relevance_details' 
			if !_rc {
				loc relevance_details = "`relevance_details'" 
			}
			else {
				loc relevance_details = "all"
			}
			putpdf paragraph
			putpdf text (`"Section relevance: `relevance_details'"'), linebreak(1) italic
			putpdf text (`"Repeat count: `repeat_count_details'"'), linebreak(2) italic 
		}

		// END REPEAT
		if "`type'" == "end repeat" {

			// define locals
			loc footing = `"${rheading`name'}"'
				
			// put in pdf 
			putpdf paragraph, font(,12) 
			putpdf text ("End Repeat: `footing'"), bold linebreak(2)
		}

		
		// NOTE
		if "`type'" == "note" {
							
			// put in pdf 
			putpdf paragraph				
			putpdf text ("Note:"), bold
			putpdf text (" `label'"), linebreak(1)
			
			capture confirm e `relevance'
			if !_rc { 
				putpdf text ("Note relevance:"), italic
				putpdf text (`" `relevance_details'"'), linebreak(2)
			} 
			
		}
		
		// TEXT, INTEGER, SELECT_ONE, SELECT_MULTI
		if "`type'" == "text" | "`type'" == "integer" | "`type'" == "select_one" |"`type'" == "select_multiple" {
		
			// update question counter 
			global question_count = $question_count + 1			
			
			// put in pdf

				// define cell containing choices, and values 
				if "`type'" == "select_one" |"`type'" == "select_multiple" {

					if `choicecount' < 10 {
						qui putpdf table cv = (`choicecount', 1), border(all, nil) memtable
						forvalues j = 1 / `choicecount' {
							loc ch: word `j' of `choices' 
							loc v: word `j' of `values' 
							putpdf table cv(`j', 1) = ("`ch'  = `v'")
						}
						putpdf table cv(1,.), addrows(1, before) 
						putpdf table cv(1,1) = ("Value label:   "), italic underline 
						putpdf table cv(1,1) = ("`list_name'") , append underline
						loc rrrows = 10 
					}
					else {
					qui putpdf table cv = (1,1), memtable
					putpdf table cv(1, 1) = (`"See value label "`list_name'""'), border(all, nil) italic
					}
				}
				
				// main question table
				qui putpdf table nt = (4,7), border(all, nil) memtable
				putpdf table nt(1,1) = ("$question_count. "), bold  
				putpdf table nt(1,1) = ("`label_details'`hint_details'"), append
				putpdf table nt(1,1), span(2, 7) 
				putpdf table nt(3,1) = ("`type'"), colspan(2) 
				putpdf table nt(4,1) = (" "), colspan(2)
				
				// insert subtable containing question name, type 
				if "`type'" == "select_one" |"`type'" == "select_multiple" {
					qui putpdf table nt(4,3) = table(cv), colspan(4) 
				}
				
				// define subtable containing name, relevance, constraints, and required status 
				if `rows' > 0 { 
					qui putpdf table auxt = (`rows', 1), border(all, nil) memtable 
					loc row = 1 
					
					capture confirm e `name' 
					if !_rc { 
						putpdf table auxt(`row', 1) = ("Var: "), underline 
						putpdf table auxt(`row', 1) = ("`name_details'"), append
						loc row = `row' + 1 
					}

					capture confirm e `group' 
					if !_rc { 
						putpdf table auxt(`row', 1) = ("Group: "), underline 
						putpdf table auxt(`row', 1) = ("`group_details'"), append
						loc row = `row' + 1 
					}

					capture confirm e `relevance' 
					if !_rc { 
						putpdf table auxt(`row', 1) = ("Relevance: "), underline 
						putpdf table auxt(`row', 1) = ("`relevance_details'"), append
						loc row = `row' + 1 
					}
											
					capture confirm e `constraint' 
					if !_rc { 
						putpdf table auxt(`row', 1) = ("Constraints: "), underline 
						putpdf table auxt(`row', 1) = ("`constraint_details'"), append
						loc row = `row' + 1 
					}
					
					capture confirm e `required' 
					if !_rc { 
						
						}
					else { 
						putpdf table auxt(`row', 1) = ("Required: "), underline 
						putpdf table auxt(`row', 1) = ("`required_details'"), append 
					}
						
				} 
				else qui putpdf table autx = (1,1), border(all, nil) memtable
				
				// Create final table
				capture confirm e `constraintmessage' 
				if !_rc loc trows = 12
				else loc trows = 10
				qui putpdf table t = (`trows', 10), border(all, nil) 
				putpdf table t(1,1) = table(nt), span(5, 7) 
				if `rows' > 0 putpdf table t(1, 8) = table(auxt), span(10, 3)    

				capture confirm e `constraintmessage' 
				if !_rc putpdf table t(11, 1) = ("Constraint Message: `constraintmessage_details'"), italic span(2, 10)
				else
			
			// update table counter 
			global table_count = $table_count + 3 
			if "`type'" == "select_one" |"`type'" == "select_multiple" global table_count = $table_count + 1
		}
		
		// CALCULATE and CALCULATE_HERE 
		if "`type'" == "calculate" | "`type'" == "calculate_here" {	
			// put in pdf 
			putpdf paragraph				
			putpdf text ("Calculate:"), bold
			putpdf text (" `name' = `calculation_details'"), linebreak(1)
			putpdf text ("`label_details'"), italic linebreak(1)
		}
 end
	