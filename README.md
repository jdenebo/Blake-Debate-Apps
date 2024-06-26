# Blake-Debate-Apps
Repo for code (primarily js and python) used for email parsers, slack/google sheets integration on google app scripts, and word doc parsers to help Blake debate sustain their model.

## Slack Integration

This is all slack integration with google apps, so it is all done with a slack app created, and then code written in google app scripts (.js). 

### Slack/Sheets Integration (slack-evidence-parser-sheets.js)

This will handle slack messages that need to be documented in the evidence mastersheet, so any message with a file attached. Will parse descriptions for topic sorting, as well as information about the message (who sent it, when they sent it, what files it included, a link to the files etc) and populate the spreadsheet accordingly. Will also support edited messages, deleted messages, and replies. 

Finally, it will remind the students when it catches mistakes at any point in the process. 

### Slack/Docs Integration (feedback-tracker.py)

This will be to automate feedback. It will pick up messages starting with a certain feedback trigger (maybe "FEEDBACK:") in each team channel and populate the proper document with that feedback so students have it all in one place. 

## Doc Parser (research-tracking-spreadsheet-verification.py)

Will parse word documents to catch cite errors, formatting errors, etc. Is integrated with slack app to immediately return feedback for students. Additionally, it will check research tracking spreadsheet to make sure that the article has not already been claimed. Will also count how many cards are there and update spreadsheet with card counts for each student to track work volume. 

## Email Parser (gmail_thread_to_docs.js)

Will sort email chain docs, filter out spam, organize blake docs into mastersheet with links to all of the relevant files. 


