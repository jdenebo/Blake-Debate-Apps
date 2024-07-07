# Blake Debate Apps
Repository for code used for email parsers, slack/google sheets integration on google app scripts, and word doc parsers to help Blake debate sustain their model. As I develop this further, I will publish the SlackApps themselves to make startup much easier. 

## Slack Integration

The slack integration portion requires the team have created a slack workspace, ideally one with the pro version (for permanent cloud storage of files). Once that has been created, follow the instructions on the slack website to create an app. We want the created app to have very little functionality, it serves exclusively to send HTTP requests to the Google Script when events occur, so event subscriptions (and the associated permissions) are all the functionality that is needed.

### Evidence Tracking Sheet (SlackEvidenceTracking.js)

This will handle slack messages that need to be documented in the evidence mastersheet, so any message with a file attached. Will parse descriptions for topic sorting, as well as information about the message (who sent it, when they sent it, what files it included, a link to the files etc) and populate the spreadsheet accordingly. Will also support edited messages, deleted messages, and replies. It will catch mistakes in this process and remind students to fix those mistakes. I reccommend making an evidence channel in slack and putting the bot in there. Messages to the channel should be in the following format: 

[Topic] - [Year] - [Argument Type] - [Argument description] 

So an example would be: Septober - 2024 - uniqueness - biden cancel border contracts. The spreadsheet would populate as follows: 

<img width="1166" alt="image" src="https://github.com/jdenebo/Blake-Debate-Apps/assets/114824938/366c0876-46da-4815-8a9b-8f0226a3bc9e">

As you can see, there are columns with useful information about the card, including year, topic, author, date the card was sent, who sent it, etc. Two important things to flag: a) it populates the spreadsheet with "YES" or "NO" to represent if there was a PDF attached to the message, and b) it populates a permalink to the evidence stored in Slack's cloud. The former is done to ensure all PDFs are available for articles to comply with NSDA evidence rules. The latter is done to give easy access to evidence at any time. If, in a year, I wanted to look for previously sent out border contract evidence for any reason, I could simply search the spreadsheet (command-F) for "border contracts" or something along those lines, and would easily find this.

### Feedback Tracker (feedback-tracker.py)

Like before, make a new bot with permissions to read messages, files, and write messages. Send events (through either slack bolt for the desktop version, or google app script for the web version) to this bot. This picks up messages starting with a certain feedback trigger (Round Feedback:) in each team channel and populate the proper document with that feedback so students have it all in one place. Again sending messages with good formatting will help populate the spreadsheet with good information. Messages should be sent with the following fomrat: 

Round Feedback: [Tournament/Round] - [Side/Order] [Team 1] vs [Side/Order] [Team 2]:

So an example would look like: Round Feedback: Practice Round 6.6.24 - 1st Aff Blake CK vs 2nd Neg Blake OR. The sheet would populate like so, with all feedback attached the line below that headline. 

<img width="1642" alt="image" src="https://github.com/jdenebo/Blake-Debate-Apps/assets/114824938/ccd5eca9-158b-4025-b154-2ea8437eb2d7">


## Doc Parser (ResearchTrackingVerification.py/ResearchTrackingVerificationCount.js)

As with before, you should make a new bot and send HTTP requests over either slack bolt or google app script (depending on desktop vs web version). Will parse word documents to catch cite errors, formatting errors, etc. Is integrated with slack app to immediately return feedback for students. Additionally, it will check a research tracking spreadsheet to make sure that the article has not already been claimed. Will also count how many cards are there and update spreadsheet with card counts for each student to track work volume. This should just be in the same evidence channel as the tracking sheet bot. 

You will also want a research tracking spreadsheet. Here (https://docs.google.com/spreadsheets/d/1Kaw3VtnnV05tBY2hsoK7ouoDOW0weQ_fVrbe_M2M4_8/edit?usp=sharing) is a template that Blake uses. The bot is set up to work on this column orientation. Students fill this out when they find an article they are cutting to ensure other students on the team do not cut the same article and waste valuable time. The bot will first check the word document sent through the evidence channel against the title column of this spreadsheet to ensure the only appearence will be "claimed by" the person who sent the message in slack. It checks this by ensuring there is only one apperance of that title in the sheet, and then checks if name in "claimed by" is the same first name associated with the slack account who sent the message in the channel. For this to work, the cite should contain the title in quotation marks, and there should be no quotation marks after the title appears in the cite (which should be no issue since the url is usually the only thing after the title in the cite). 

Additionally, cites should end with the initials of the person who cut the card. For example, mine end with )JDE. This bot will also assign the number of cards in the document to those initials. It will query a spreadsheet that keeps track of article and card counters per initial and update those values to count the number of cards cut by each kid. This spreadsheet needs to be created and maintained by you. I have it set up with the names and initials in first two columns, and then column sets for each topic to track articles/card counts. I will upload a template of this at some point in the future. 

## Email Parser (GmailThreadToDocs.js)

This does not require a bot, and simply is done through google app script. The only thing needed is to edit the script to check for whatever label you send email chains to. We have a google group that is put on every chain and I have it redirect emails to my debate email, and sort them into a label while starring them. The script will look into threads under that label that are starred and process them. Simply make a google drive folder and it will load every thread into a unique folder within it. I set a trigger to update it every 10 minutes. 

