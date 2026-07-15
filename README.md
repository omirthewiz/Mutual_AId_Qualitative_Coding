# Transcript Coding Pipeline

Tools for turning workshop transcripts into an Excel workbook you can code in, for the Mutual AI(d) Lab.

You give it a transcript. It cleans the transcript, hides real names, and builds an Excel file where each workshop gets its own tab, set up so each person can code on their own and then everyone's codes get combined.

## Credit

The transcript format and the cleaning step build on Wentao Guo's clean-transcript (github.com/oatnewguo/clean-transcript). zoom_parse.py and transcript.py were written/additions for this repo.

## What this repo includes

This is a starter template. It has the scripts and an example codebook, but no real data. You bring your own transcripts and your own codebook.

## How it works

There are two ways to start, depending on what your transcript looks like.

1. If you already have a VTT file (like a Zoom download or a Whisper output), you can go straight to clean.py. It works fine on its own. You don't need to do anything first.

2. If all you have is a Zoom transcript you copy-pasted (the messy text where each name shows up twice with timestamps mixed in), run zoom_parse.py first. It turns that paste into a clean VTT that clean.py can read.

So the order is:

1. zoom_parse.py (only if you copy-pasted from Zoom. Skip it if you already have a VTT.)
2. clean.py (cleans the text and hides names)
3. transcript.py (builds the Excel file)

## Setup

You need Python 3 and the packages in requirements.txt.

```
pip install -r requirements.txt
```

## The scripts

### zoom_parse.py (only for Zoom copy-paste)

Turns a copy-pasted Zoom transcript into a clean "Name: text" VTT. It figures out the speakers on its own, so no names are typed into the code. It keeps timestamps, joins lines from the same speaker, and prints the list of speakers it found so you know who to hide in the next step. It does not hide names itself.

Only use this when you don't have a VTT and had to paste the transcript in by hand.

```
python3 zoom_parse.py raw.txt -o workshop3.vtt
```

### clean.py

Fixes common transcription mistakes and hides real names by swapping them for IDs. Run it once for each name, using -n for the real name and -a for the ID. Do this for nicknames too, and change any facilitator's name to "Interviewer".

```
python3 clean.py workshop3.vtt -n "Real Name" -a "P1"
```

You can change what text gets fixed by editing replacements.py. Instructions are in that file.

### transcript.py

Builds the Excel file. Each transcript becomes its own tab in coding.xlsx (it makes the file if it isn't there yet).

```
python3 transcript.py workshop3.vtt
```

What each tab looks like:

1. The columns are Speaker (A), Timestamp (B), Utterance (C), then coding columns: Dr.Ming (D/E/F), Omir (G/H/I), Hawi (J/K/L), and Combined (M/N/O/P) for the final agreed codes.
2. Any "Researcher #1/#2" style name becomes just "Interviewer".
3. Long participant turns (more than 5 sentences) get split into rows of about 5 sentences each, so each piece can be coded on its own. Interviewer turns are left whole.
4. The list of participant IDs goes in cell R1 ("Participants: ..."), not the tab name, because a tab name can only hold 31 characters.
5. The code cells have dropdown menus that pull from the Codes sheet. You can still type in a new code that isn't in the list yet.
6. If you type a code that isn't in the codebook, the cell turns yellow. That's the sign someone added a possible new code. It goes back to normal once you add that code to the Codes sheet.

It also makes helper sheets:

1. Codes is your codebook: the code, a description, and how many times each code was used.
2. GetCodes gathers every code that was used, pulling only from the Combined columns (M-P), with the quote, speaker, and timestamp next to it.
3. NewCodes lists any Combined-column code that isn't in the codebook yet, so the team can review it.

## Step-by-step example

Starting from a Zoom copy-paste:

1. Paste the Zoom transcript into a text file, like ws3_raw.txt.
2. Run `python3 zoom_parse.py ws3_raw.txt -o workshop3.vtt`. Look at the speakers it prints.
3. For each real name, run `python3 clean.py workshop3.vtt -n "Real Name" -a "P1"`. Change any facilitator to "Interviewer".
4. Run `python3 transcript.py workshop3.vtt`. This adds a "workshop3" tab to coding.xlsx.
5. Open coding.xlsx and turn on the two helper sheets (see "Turning on GetCodes and NewCodes" below).

If you already have a VTT, skip steps 1 and 2 and start at step 3 (or step 4 if names are already hidden).

## codes.txt vs the Codes sheet

This trips people up, so read this part.

codes.txt and an existing Excel file are not connected. codes.txt is only used when transcript.py makes a brand-new Excel file. Editing codes.txt does nothing to a file that already exists.

So:

1. To change the codebook in a file you already have, edit the Codes sheet in Excel. Just type a new code in the next empty row. The dropdowns, the counts, the yellow highlighting, and NewCodes all update on their own, because their formulas already reach down to row 200.
2. If you add codes to an existing file, also add them to codes.txt, so the next brand-new file starts with the current codebook.

Short version: codes.txt is the starting point for new files. The Codes sheet is the live list.

## Why GetCodes only uses Combined

Each person codes in their own columns. A code isn't final until it's moved into the Combined columns (M-P). GetCodes only reads Combined, so a workshop won't show up there until its Combined columns are filled in.

## Turning on GetCodes and NewCodes

transcript.py can't write these two formulas straight into Excel, so it saves them to getcodes_formula.txt and newcodes_formula.txt instead. To turn on a sheet, open the matching text file, copy everything, and paste it into cell A2 of that sheet (GetCodes or NewCodes) in Excel.

## If something goes wrong

1. When openpyxl opens and saves a file, it removes the dropdowns and the yellow highlighting. It warns you with "extension is not supported and will be removed". So running transcript.py again on a file you already have, or opening and saving in some versions of Excel, can wipe them out. You may need to add them back by hand in Excel, which actually holds up better.

2. The GetCodes and NewCodes formulas sometimes don't refresh. Re-typing the formula in the top cell, or deleting the sheet and making it again, fixes it.

3. If a cell that should turn yellow stays blank, the highlight color needs to be written as FFFFFF00. If it starts with "00" instead, the color is invisible.

## One thing to know

transcript.py only puts the transcript into the Excel file. It does not clean the text or hide names. That all happens in clean.py. If your transcripts already have IDs instead of names and are already cleaned up, you can skip clean.py. Just know that if you skip it, nothing fixes typos or removes "um" and "uh", so make sure the transcript is already in good shape.
