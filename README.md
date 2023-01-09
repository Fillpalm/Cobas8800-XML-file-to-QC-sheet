### XML file parsing script

Who: Author -> Phill Palmer 
What: A script to automatically convert the XML file xported from the cobas RUI to usable data
When: last updated January 9th, 2023
Where: Cherry Hill NJ molecular virology lab
Why: To automate QC data upload after runs on the cobas 8800. also to extract Ct information
Required: spefically formatted QC sheets excel file. "new XML files" and "old XML files" folders, .bat file with PATHs set. only on Windows currently.

How: Currently, the script only functions to update our specfic cobas 8800 QC file, for specified tests. 

	1) Export the XML run file from the cobas RUI 
		a) select batch from Routine >control batch 
		b) change "Not released" to "All"
		c) select all samples (check the box at the top)
		d) hit "export"
		e) name the file
			- file must start with "b###" where ### is the control batch ID
			- TEST NAME MUST BE INCLUDED IN FILE NAME! ex: hiv
		f) on the remote user interface (RUI) go to Administration >File management >exports >select file >download
		g) switch the file type to text and add ".xml" to the end of the file
		h) save to the folder "new XML files"
	2) in the QC sheet, add manual information in the tan columns. Only "control batch #" is required 
	3) exit out of the file and double click the "XML update script" icon on the desktop
		- processed XML files will be moved to the "old XML files" folder
		- any errors will pop up in the command window
		- the window automatically closes after 10 seconds


the script can be easily modfied to produce two excel files:
	- one with reagent information
 	- one with results information
