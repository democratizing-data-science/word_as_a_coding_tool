# word_as_a_coding_tool
This macro enables the use of word as a coding device, substituting NVivo, Atlas.ti, RQDA, etc. 

# Make sure you install the macro in PC or Mac depending on your operating system

The macro exports all comments (i.e., codes) and their content (i.e., quotes) to a new Excel document in a table form.

This table can then be used with MOVIE to analyze the temporal evolution of those codes.

***IMPORTANT*** If the Word document (transcript) contains multiple participants, like in focus groups, add the participant ID to each comment followed by ":" as follows:
UNIT_ID_i:Code_j. Each comment needs to start with the actor ID information, which may be more than one in cases with multiple participants providing information.

After all coding is completed, you will need to select the column called "UNIT_ID_i: Code_j" in excel and select Data, Text to Columns using ":" (without quotations) as the delimiter in options. After this, you will have two columns, one with actors ID and another with codes IDs. Even if the latter column is called code, it can be used for communication exchanges in the form Actor_ID_i to Actor_ID_j. 

Recall that *All these column names are optional* for MOVIE was programmed to work regardless of the column names selected. The only requirement is that the relationships are in the form, sender to receiver or that the column UNIT_ID sends a connection to the column containing codes or being the target of a message.

The analyses of actor to actor communication exchanges may be achieved by selecting the text in the word document and adding a comment in the form: Actor_id_i:Actor_id_j. The process of separating these relationships using Data, Text to Columns using ":" (without quotations) as the delimiter is the same as detailed in the example.

***Instructions to USE***
To use, open word in windows or mac.

Go to view, 

Select macro, 

Add a temporal name if no macro exists, then select create in windows (or click the plus sign in mac)

Once visual basic for applications (VBA) is opened, select the code provided here and replace all info in the new macro in VBA word. 

Save and close.

When you open a word document with comments, select view, macro, the macro called "Word_as_a_coding_device" and hit run. All codes will be exported.

***You only need to do these steps once per machine***
