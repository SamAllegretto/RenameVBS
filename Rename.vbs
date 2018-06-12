dim fso
dim folderspec
dim answer
dim Hull
dim comp

Hull = InputBox("Enter Hull Number")
comp = InputBox("Enter Compartment Name")

set fso = CreateObject("Scripting.FileSystemObject")
folderspec = fso.GetParentFolderName(WScript.ScriptFullName)

set f = fso.GetFolder(folderspec)

for each f1 in f.files

	Oldfile = f1.name
	M_file = f1.name
	'MsgBox(fso.GetExtensionName(f1.Name))
	
	If inStr(fso.GetExtensionName(f1.Name), "JPG") > 0 or inStr(fso.GetExtensionName(f1.Name), "JPEG") > 0 or inStr(fso.GetExtensionName(f1.Name), "jpg") > 0 or inStr(fso.GetExtensionName(f1.Name), "jpeg") > 0 Then
		
		if inStr(Oldfile, "-") > 0 then
			'answer=MsgBox(Len(Oldfile) - instr(f1.name,"-"),65,"index")
			Oldfile = Right(Oldfile,Len(Oldfile) - instr(f1.name,"-"))
			Oldfile = Right(Oldfile,Len(Oldfile) - instr(Oldfile,"-"))
			
		end if
		
		if Hull = "" or comp = "" then
			Newfile = Oldfile
		else
			Newfile = Hull & "-" & comp & "-" & Oldfile
			fso.MoveFile M_file, Newfile 
		end if
		
		'MsgBox(Newfile)
		
	end if
	
next