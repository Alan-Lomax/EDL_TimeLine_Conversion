# Usage Case

When posting videos to YouTube or another Website / Wiki it is desirable to include a list of chapter markers where
the user reading about the video can just click on the marker and jump to the corresponding section of the video.

YouTube can do this if you paste an appropriately formatted text into the description when you publsih the video.
Websites and Wikis can also do this but the format is substantially different. 
Creating this formatted text is tedious and subject to errors.

# EDL_TimeLine_Conversion
## An Excel Automation Tool

In video editing software it is common to be able to save a Edit Data List (EDL file) which has the raw data needed (but not in the right format)
This spreadsheet tool will import such an EDL file and transform it into two different text formats.
One is a text format that can be pasted directly into the YouTube description box (when publisinbg the video)
The second is n the form of a markup code snippet (with the same chapter content) that can be pasted into a Wiki (for example the MERG knowledge base.)

The final step is for the user to copy either of these texts and paste it in the destination.

## How it Works
Using a button on the instructions page a template is copied to a working sheet
On the working sheet another button invkes a VBA macro 
Under the control of that Excel VBA Macro
   A windows dialog box is used to locate an EDL file of your choice
   The file is handed off to an extensive PowerQuery script that formats and prepares the EDL file
   The output text is dropped into an Table on an Excel sheet dedictated for that purpose.
In the Working copy of the main Excel Spreadhseet
   Extensive text processing is done to take the processed EDL table and surround it with appropriate 'fixed text'
   
