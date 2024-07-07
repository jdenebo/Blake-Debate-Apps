from slack_bolt import App
from slack_bolt.adapter.socket_mode import SocketModeHandler
import pygsheets
from datetime import date, datetime



APP_TOKEN = #SLACK APP TOKEN
BOT_TOKEN = #SLACK BOT TOKEN
app = App(token=BOT_TOKEN)

gc = pygsheets.authorize(service_file='creds.json')
sh = gc.open(#WHATEVER SPREADSHEET NAME YOU ARE USING)

@app.event("message")
def feedback_upload(event, say, client):
  #ensures the event it picked up has the correct dictionary key
    if "text" in event.keys():
        msg_text = event["text"]
      #checks that it has the feedback trigger with a bunch of formatting stripped, in this case, the trigger is "Round Feedback:"
      #in addition to a trigger, i have implemented this with the following header in mind (to collect relevant tourney info) to go above a newline for each new round feedback
          # Round Feedback: Glenbrooks R1 - 1st Aff Blake OP vs 2nd Neg Blake EW

        if msg_text.strip().lower().startswith("round feedback:") or msg_text.strip().lower().startswith("*round feedback:"):        
            coach = client.users_profile_get(user=event["user"])
            coach_name = coach["profile"]["real_name"]

          #derives the team code from the channel this is sent in
            channelslist = client.conversations_list(exclude_archived=True, types=["private_channel"])
            channel_id = event["channel"]
            team_code = None
            for c in channelslist["channels"]:
                if c['id'] == channel_id:
                    if c['name'].upper().startswith("BLAKE"):
                        team_code = c['name'][5:]
                    else:
                        team_code = c['name']
            #the spreadsheet has individual sheets that are named after each team codes, here i am just selecting the corresponding one
            wks = sh.worksheet('title', team_code.upper())

            #adds a new row to the top
            wks.insert_rows(row=1, number=1, values=[None, None, None, None, None, None, None])

            #date value to be added to a new row, indicating a new round with new feedback
            timestamp = float(event["ts"])
            today = date.today()

            #updates row with date information
            wks.update_row(index=2, values =[None, None, today.strftime("%m/%d/%Y"), None, None, None, None])


            #gets body of feedback, coach who sent it, round and tournament info as sent in the header for the round (see above explanation)
            title = msg_text.split('\n', 1)[0]
            body = msg_text.split('\n', 1)[1:]
            body_joined = "\n".join(body)
            wks.update_row(index=2, values =[None, None, None, None, None, body_joined, None])
            info = title.split(":")[1]
            wks.update_row(index=2, values =[info, None, None, None, None, body_joined, None])

            wks.update_row(index=2, values =[info, None, None, None, None, body_joined, coach_name])

            tournament = info.split("-")[0]
            wks.update_row(index=2, values =[info, tournament, None, None, None, body_joined, coach_name])

            sides = None
            if "-" in info:
                sides = info.split("-")[1]
            else:
                say("Please make sure to separate tournament/round information from sides/teams with a hyphen like so: \nRound Feedback: Glenbrooks R1 - 1st Aff Blake OR vs 2nd Neg Fairmont CK")

          #here, i am parsing the header to find the team names/sides so i can populate the spreadsheet accordingly
            team1 = None
            team2 = None
            
            if sides:
                if "vs." in sides:
                    team1 = sides.split("vs.")[0].lower()
                    team2 = sides.split("vs.")[1].lower()
                elif "vs" in sides:
                    team1 = sides.split("vs")[0].lower()
                    team2 = sides.split("vs")[1].lower()
                elif "and" in sides:
                    team1 = sides.split("and")[0].lower()
                    team2 = sides.split('and')[1].lower()
                elif "v." in sides:
                    team1 = sides.split("v.")[0].lower()
                    team2 = sides.split('v.')[1].lower()
                
                my_team = None
                other_team = None
                if team1 and team2:
                  #addressing case of when blake does a practice round so both teams are blake
                    if "blake" in team1 and "blake" in team2:
                        if team_code.lower() in team1:
                            my_team = team1
                            other_team = team2
                        else:
                            my_team = team2
                            other_team = team1
                    elif "blake" in team1:
                        my_team = team1
                        other_team = team2
                    elif "blake" in team2:
                        my_team = team2
                        other_team = team1
                    else:
                      #just a catch all so that some information does get populated in any case if no one is doing the header properly
                        if team_code.lower() in team1:
                            my_team = team1
                            other_team = team2
                        else:
                            my_team = team2
                            other_team = team1
                    
            wks.update_row(index=2, values =[info, tournament, None, other_team.strip(), my_team.strip(), body_joined, coach_name])
            ###i have it so that it updates a row along the way instead of one batch add at the end to better deal with potential errors

if __name__ == "__main__":
    SocketModeHandler(app, APP_TOKEN).start()
