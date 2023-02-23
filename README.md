# Usage Case

When posting videos to YouTube or another Website / Wiki it is desirable to include a list of chapter markers where
the viewer can just click on the chapter marker and jump right to the corresponding section of the video.

YouTube can do this if you paste an appropriately formatted text into the description when you publsih the video.
Websites and Wikis can also do this but the format is substantially different. Many Wiki's can also do this if the appropriate 'markup' commands are used. 
Creating these snippets by hand is tedious and subject to errors.

# EDL_TimeLine_Conversion
## An Excel Automation Tool

In video editing software it is common to be able to use markers on the timelne to indicate chapters.
These markers can be saved as an Edit Data List (EDL file) which has the raw data needed for our automation..
This spreadsheet tool will import the EDL file and transform it into two different output text formats.
One is a text format that can be pasted directly into the YouTube description box (when publisinbg the video)
The second is in the form of a markup code snippet (with the same chapter content) that can be pasted into a Wiki (for example the MERG knowledge base.)

After using this tool the final step is for the user to copy one or both of these text outputs and paste them into the destination.

## How it Works
Using a button on the instructions page a template is copied to a working sheet
On the working sheet a second button invokes VBA code which:
   Opens A windows dialog box to locate an EDL file of your choice
   The file is handed off to an extensive PowerQuery script that reads the EDL file (The EDL file is unchanged)
   The Script output text is dropped into a new Query Table on the EDL processing sheet (normally hidden).
In the Working copy of the main Excel Spreadhseet
   There are a few cells which must be manually entered:
      Link to the YouTube video
      Titles and Dates of the Video
      A One Line Description amd a Longer Description.
   Extensive text processing is done to take the processed EDL data and surround it with appropriate 'fixed text' and these values.
   Two output blocks show the final texts.
   Once happy with the results a Third and Final button will convert all formulae to values - so the Working copy is preserved. 
   (The working copy can be manually edited - but all of the formuale are removed.)
   
   ### Note:
   If the working sheet was not 'frozen' in this way ..
   Any subsequent processing of a different EDL file will update the EDL_Processing sheet.
   This would cause all working sheets that had not been frozem to update - which is not likely the intent. 
   
