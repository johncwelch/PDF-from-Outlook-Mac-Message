--Note that this script creates temp HTML and text folders in your home directory. It doesn't delete them. 
--If that is not behavior you like, it is trivial to change it.

global theTempFolder --alias to folder with temp files
global theTempFilePath --path used for a few things
global theMessageTitle --subject of email message

property theTempFileList : {} --list of files passed to Acrobat
property badSelection : false --did someone select something silly?

set theSystemInfo to system info --get the base system info
set theHomeDir to home directory of theSystemInfo as text --get the current user home directory
set theDocumentsFolder to theHomeDir & "Documents:" --we get the path to the home directory documents folder,
--because it's where we store the source for the message.
try
	set theTempFolder to (theDocumentsFolder & "outlook2pdf:" as alias) --we're looking for our temp folder
	
on error theErrorMessage number theErrorNumber --test the error
	if theErrorNumber = -43 then ---43 is the error number for folder doesn't exist, so that's the only time we want to create it
		--you can add code for other errors if you like
		tell application "Finder"
			set theTempFolder to ((make new folder at (theDocumentsFolder as alias) with properties {name:"outlook2pdf"}) as alias) --make the folder in the current user's documents directory
		end tell
	end if
end try

tell application "Microsoft Outlook"
	set theMessageList to selected objects --new for Outlook
	repeat with x in theMessageList --traverse the list
		set theMessageProps to properties of x --get message props
		if class of theMessageProps is not incoming message then --something is very, very wrong
			set badSelection to true
			exit repeat --we're done here, no use even trying the rest of the list, something is way out of whack for this script to handle
		end if
		set badSelection to false --it's okay
		set theMessageHasHTML to has html of theMessageProps --does outlook think this is HTML? Yes, I know, less than perfect way to do this
		--but it's fast and reliable enough, and if Outlook doesn't know the message has HTML, then you probably won't just by looking at it
		set theMessageTitle to subject of theMessageProps --this will become the file name
		set theMessageTitle to my hasIllegalCharsInSubject(theMessageTitle) --we only care about removing slashes and colons. the rest can stay
		--since this only runs on a Mac, we don't care about Windows issues.
		set theMessageContent to content of theMessageProps
		if theMessageHasHTML then --HTML email
			set theTempFilePath to (theTempFolder as text) & theMessageTitle & ".html" --we'll save it as an HTML file. This is so Acrobat can open it later.
			set theFileHandle to open for access theTempFilePath with write permission --open up the file we'll create. 
			write theMessageContent to theFileHandle as «class utf8» --write the content to disk as UTF-8
			--avoids unwanted character conversion
			close access theFileHandle --close the file
			set the end of theTempFileList to theTempFilePath --add this to the list of files for Acrobat to open
		else if not theMessageHasHTML then --it's not HTML, must be plain text.
			set theTempFilePath to (theTempFolder as text) & theMessageTitle & ".txt" --save it as a text file
			set theFileHandle to open for access theTempFilePath with write permission
			write theMessageContent to theFileHandle as «class utf8» --write the content to disk as UTF-8
			close access theFileHandle
			set the end of theTempFileList to theTempFilePath
		else
			display dialog "couldn't determine if this message has HTML or not. Skipping" --this should never happen, but just in case, we'll
			--process the rest of the list anyway. No sense in hosing it all for one problem child. 
		end if
	end repeat
end tell

if badSelection then
	display dialog "This script ONLY handles messages, not folders or anything else. Please verify you only have a message selected." giving up after 60
else
	display dialog "converting the messages to PDF. This may take a while, so be calm" buttons "Okay" giving up after 30
	repeat with x in theTempFileList
		set theTempFilePath to (contents of x as alias)
		tell application "Adobe Acrobat" --if you have question marks in the file name, Acro 10.X may bugger this all for a lark until Adobe fixes this bug.
			--Acrobat DC seems to work fine.
			set thePDFConvert to open theTempFilePath --open the file in Acrobat. 
		end tell
		
		(*tell application "PDFpenPro.app" --if you want to use PDFPenPro instead of Acrobat, uncomment this block
		--and comment out the Adobe Acrobat tell block.
			make new document with data theTempFilePath
		end tell*)
		
	end repeat
	
end if

on hasIllegalCharsInSubject(theSubject)
	set theTest to offset of ":" in theSubject
	if theTest ≠ 0 then
		set oldDelims to AppleScript's text item delimiters
		set AppleScript's text item delimiters to ":"
		set theBadCharList to (every text item of theSubject)
		set AppleScript's text item delimiters to "-"
		set theSubject to theBadCharList as text
		set AppleScript's text item delimiters to oldDelims
	end if
	
	set theTest to offset of "/" in theSubject
	
	if theTest ≠ 0 then
		set oldDelims to AppleScript's text item delimiters
		set AppleScript's text item delimiters to "/"
		set theBadCharList to (every text item of theSubject)
		set AppleScript's text item delimiters to "-"
		set theSubject to theBadCharList as text
		set AppleScript's text item delimiters to oldDelims
	end if
	
	return theSubject
end hasIllegalCharsInSubject
