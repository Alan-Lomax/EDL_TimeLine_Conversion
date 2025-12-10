# Usage Case

When posting videos to YouTube or another Website / Wiki it is desirable to include a list of chapter markers where the viewer can just click on the chapter marker and jump right to the corresponding section of the video. This excel does simple text manipulation to produce the list of chapter markers automatically.


# Work Flow
o  Edit your video in DaVinci resolve or similar. Create chapter markers at significant points in the video. (not more than one every 5 minutes

o  Render your video for publishing to YouTube

o  From the Timeline Export the markers to an EDL file

o  Start this worksheet. On the Instructions tab is a button to create a new worksheet. (it copies a template sheet as a starting point)

o  In YouTube upload your video. 

o  Copy the Link and paste it into the spreadsheet where indicated. Note: Very early in the upload process YouTube presents you with a link that will be valid for your video (even though it may not have finished uploading yet)

o  Click the button “Import EDL”. A standard windows dialog will open so you can find the file and click on it. Once done an excel power query will read and parse the EDL file and drop the resulting data into your worksheet.

o  Fill in any remaining data as desired (yellow boxes at top of sheet). Date and duration of the video for example.

o  At the bottom of the sheet are two output areas. One for YouTube and one for the MERG Wiki.

o  Copy the YouTube data from the Excel and past it into the description box for your YouTube video. (Which may still be uploading)

o  Copy the MERG Wiki Knowledgebase output for pasting into the WIKI as desired.



# EDL_TimeLine_Conversion
## An Excel Automation Tool

In video editing software it is common to be able to use markers on the timeline to indicate different chapters. These markers can then be saved as an Edit Data List (EDL file). This spreadsheet tool imports the EDL file and transforms it into two different output text formats.

Format One: Can be pasted directly into the YouTube description box (when you are first publishing the video). 

Format Two: Is a markup code snippet that can be pasted into a Wiki while editing a page. (for example, the MERG knowledge base.)

The excel produces both of these formats at the same time from the imported EDL file. 

The final step is for the user to copy  one or both of these text outputs and pastes them into the appropriate destination.


## Benefits
- Automatic chapter markings from the EDL list
- Consistent error free code snippets
- Efficiency: The total work time is less than 5 minutes 

## How it Works
Using a button on the instructions page a template worksheet is copied over to a working sheet. (Tab Color is Yellow)

On the working sheet just created a second button:
   + Opens A windows dialog box to locate an EDL file of your choice
   + The file is handed off to an extensive Power Query script that reads the EDL file (The EDL file itself is not altered)
   + The Script output text is dropped into a new Query Table on the EDL processing sheet (which is hidden).

In the Working copy of the main Excel Spreadsheet
   + There are a few cells which must be manually entered:
     +   The Link to the YouTube video (available as soon as uploading to YT has begun)
     +   Titles and Dates of the video
     +   Duration of the video in minutes
     +   A One Line Description and a Longer Description.

     + Extensive text manipulation is done to surround the processed 
       EDL data with fixed text and the above one-time values.
   + Two output blocks contain the final combined text.
   + These output texts can be copied to the destination (YouTube or Wiki)

Once happy with the results a Third and Final button will convert all formulae to values - so the Working copy is preserved. (Tab Color changes to Green) The working copy can still be manually edited and copied to the destination but all of the formulae are removed.
   

### Note 1:
If the working sheet was not 'frozen' any subsequent processing of a different EDL file will update the EDL_Processing sheet.
This would cause all working sheets that had not been frozen to update - which is not likely the intent. 

### Note 2:
In Davinci resolve the default behaviour has the timeline start at 01 hrs. For Chapter markers to work properly this must be reconfigured so the start timecode is 00 hrs.
