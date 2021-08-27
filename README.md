# word_as_a_coding_tool
This macro enables the use of word as a coding device, substituting NVivo, Atlas.ti, RQDA, etc. 

The macro exports all comments (i.e., codes) and their content (i.e., quotes) to a new Excel document in a table form.

This table can then be used with MOVIE to analyze the temporal evolution of those codes.

***IMPORTANT*** If the Word document (transcript) contains multiple participants, like in focus groups, add the participant ID to each comment followed by ":" as follows:
UNIT_ID_i: Code_j
Only when working with group transcripts, you will need to select the column called "UNIT_ID_i: Code_j (if group transcript)" and select Text to Columns using : as the delimiter in options. After this, you may delete "(if group transcript)" in the resulting column called "(Code_j (if group transcript)." The same appies to the column called "UNIT_ID (if single case, e.g., interview transcript)", that is you can delete the "(if single case, e.g., interview transcript)" portion. Finally, if analysing individual files you can change the column name "UNIT_ID_i: Code_j (if group transcript)" to "Code_j." All these changes to name are optional for MOVIE was programmed to work regardless of the genering column names selected. The only requirement is that the column UNIT_ID sends a connection to the column containing codes.

The analyses of actor to actor communication exchanges need manual data preparation as explained in the example.

***USE***
To use, open word in windows or mac.

Go to view, 

Select macro, 

Add a temporal name if no macro exists, then select create in windows (or click the plus sign in mac)

Once visual basic for applications (VBA) is opened, select the code provided here and replace all info in the new macro in VBA word. 

Save and close.

When you open a word document with comments, select view, macro, the macro called "Word_as_a_coding_device" and hit run. All codes will be exported.

***You only need to do these steps once per machine***
