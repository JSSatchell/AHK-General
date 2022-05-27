#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.
#SingleInstance force

movie:=[]
comingAttraction:=[]
oldMovie:=[]
r=
c=

;;;;;Automatic Update;;;;;
MsgBox, 4, Update method, Press 'Yes' to update from current CSV.`nPress 'No' to enter new upcoming films.
IfMsgBox, Yes
{
	CSV_Load("Filmlings_Full Lineup.csv","films")
	Gosub, LoadMovies
	Gosub, LoadComingAttractions
	Gosub, PopulateMovieFiles
	Gosub, PopulateComingAttractionFiles
	Gosub, Confirmation

}

;;;;;Manual Entry;;;;;
IfMsgBox, No
{
	r:=1
	while (r<=4) {
		c:=1
		;;;;;Coming Attraction 1;;;;;
		switch r 
		{
			case 1:
				w = first
				wCap = First
			case 2:
				w = second
				wCap = Second
			case 3:
				w = third
				wCap = Third
			case 4:
				w = fourth
				wCap = Fourth
		}

		MsgBox,, % wCap . " Film", % "Please enter information for the " . w . " upcoming film."

		while (c<=5){
			switch c
			{
				case 1:
					w2 = title
					w2Cap = Title
				case 2:
					w2 = year
					w2Cap = Year
				case 3:
					w2 = IMDb URL
					w2Cap = IMDb
				case 4:
					w2 = trailer URL
					w2Cap = Trailer
				case 5:
					w2 = Wikipedia URL
					w2Cap = Wiki
			}

			Loop
			{
				InputBox, val, % w2Cap . " " . r, % "Enter the " . w2 . " of the " . w . " upcoming film."
				comingAttraction[r,c] := val
				if ErrorLevel
				{
				MsgBox, 4, Continue?, Are you sure? If you exit all input will be lost.
				IfMsgBox, No
					Continue
				IfMsgBox, Yes
					ExitApp
				}
				else Break
			}

			c++
		}

		r++
	}

	r:=1
	while(r<=4){
		c:=1
		FileAppend, `n, Filmlings_Full Lineup.csv
		while(c<=5){
			a:=comingAttraction[r,c]
			FileAppend, % c=5 ? a : a . ",", Filmlings_Full Lineup.csv
			c++
		}
		r++		
	}

	CSV_Load("Filmlings_Full Lineup.csv","films")

	r:=1
	while(r<=4){
		oldMovie[r]:=CSV_ReadCell("films", r+1,1)
		r++
	}

	MsgBox, 1, Delete Alert, % "You are about to delete:`n`n" . oldMovie[1] . "`n" oldMovie[2] . "`n" . oldMovie[3] . "`n" oldMovie[4]
	IfMsgBox, Cancel
	{
		MsgBox, 4, Continue?, Would you like to continue? If not, all input will be lost.
		IfMsgBox, No
			ExitApp
	}

	i:=1
	while(i<=4)
	{
		CSV_DeleteRow("films", 2)
		i++
	}

	Gosub, LoadMovies
	Gosub, PopulateMovieFiles
	Gosub, PopulateComingAttractionFiles
	Gosub, Confirmation

	CSV_Save("Filmlings_Full Lineup.csv","films")
	; When using this CSV library the script will remain persisent (remain active in memory)
	; so you need close the script when you are done to free the memory
}
ExitApp

LoadMovies:
	r:=1
	while (r<=4){
		c:=1
		while (c<=5){
			movie[r,c]:=CSV_ReadCell("films", r+1,c)
			c++
		}
		r++
	}
return

LoadComingAttractions:
	r:=1
	while (r<=4){
		c:=1
		while (c<=5){
			comingAttraction[r,c]:=CSV_ReadCell("films", r+5,c)
			c++
		}
		r++
	}
return

PopulateMovieFiles:
	r:=1
	while(r<=4){
		c:=1
		while(c<=5){
			l=
			a=
			switch c
			{
				case 1: l=a
				case 2: l=b
				case 3: l=c
				case 4: l=d
				case 5: l=e
			}
			a := movie[r,c]
			FileDelete, M%r%\M%r%%l%.csv
			FileAppend, %a%, M%r%\M%r%%l%.csv
			c++
		}
		r++
	}
return

PopulateComingAttractionFiles:
	r:=1
	while(r<=4){
		c:=1
		while(c<=5){
			l=
			a=
			switch c
			{
				case 1: l=a
				case 2: l=b
				case 3: l=c
				case 4: l=d
				case 5: l=e
			}
			a := comingAttraction[r,c]
			FileDelete, CA%r%\CA%r%%l%.csv
			FileAppend, %a%, CA%r%\CA%r%%l%.csv
			c++
		}
		r++
	}
return

Confirmation:
	MsgBox % "Here's the new lineup:`n`n" . movie[1,1] . " (" . movie[1,2] . ")`n" . movie[2,1] . " (" . movie[2,2] . ")`n" . movie[3,1] . " (" . movie[3,2] . ")`n" . movie[4,1] . " (" . movie[4,2] . ")"

	MsgBox % "And here's what's coming up:`n`n" . comingAttraction[1,1] . " (" . comingAttraction[1,2] . ")`n" . comingAttraction[2,1] . " (" . comingAttraction[2,2] . ")`n" . comingAttraction[3,1] . " (" . comingAttraction[3,2] . ")`n" . comingAttraction[4,1] . " (" . comingAttraction[4,2] . ")"
return
