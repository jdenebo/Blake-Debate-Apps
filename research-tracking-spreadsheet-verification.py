import requests
import docx2txt
import pandas as pd
from slack_bolt import App
from slack_bolt.adapter.socket_mode import SocketModeHandler


APP_TOKEN = APP_TOKEN
#insert your respective tokens
BOT_TOKEN = BOT_TOKEN


app = App(token=BOT_TOKEN)


sheet_id = #whatever sheet ID you want
sheet_name = #whatever sheet you want
url = f"https://docs.google.com/spreadsheets/d/{sheet_id}/gviz/tq?tqx=out:csv&sheet={sheet_name}"
sheet_data = pd.read_csv(url)


def get_author_and_title(doc):
    """
    Given the title of a docx document that contains the recently uploaded cut card, this function grabs the title and author
    of the article and returns them both in a tuple.
    """

    text = docx2txt.process(doc)
    p_sections = text.splitlines()

    # grabbing cite text, finding it by checking for conditions that are exclusive to the cite (parens, DOA, urls etc)
    for i in range(10):
        cite_text = p_sections[i]
        if ("(" not in cite_text) or ((chr(8220) not in cite_text) and (chr(34) not in cite_text)) or ("http" not in cite_text):
            continue
        else:
            break

    # finding and grabbing the title from the cite text

    # if used curly quotes
    if chr(8220) in cite_text:
        left_quote_idx = cite_text.find(chr(8220))
        right_quote_idx =cite_text.find(chr(8221))
        title = cite_text[left_quote_idx+1 : right_quote_idx]

    # if used straight quotes
    elif chr(34) in cite_text:
         split_cite = cite_text.split(chr(34))
         title = split_cite[1]

    # finding author
    author = cite_text.split()[0]

    return (title, author)

def check_spreadsheet_title(word):
    """
    Checks spreadsheet for the article title, if it is in there, it returns the row that contains it. If its is not, it returns an empty dataframe.
    """
    filtered_df = sheet_data[sheet_data.iloc[:, 3].str.contains(word).fillna(False)]
    return filtered_df

def check_spreadsheet_author(word):
    """
    Checks spreadsheet for article author, if it is in there, it returns the row that contains it. If its is not, it returns an empty dataframe.
    """
    filtered_df = sheet_data[sheet_data.iloc[:, 0].str.contains(word).fillna(False)]
    if filtered_df.empty:
        return 0
    else:
        return filtered_df

@app.event({"type" : "message", "subtype" : "file_share"})
def get_file(event, say, client):
    if "thread_ts" not in event.keys():
        username = client.users_profile_get(user=event["user"])
        first_name = username["profile"]["real_name"].split(" ")[0]
        docx_counter = 0
        for file in event["files"]:
            if file["filetype"] == "docx":
                docx_counter += 1
                url = file["url_private"]
                file_name = event["files"][0]["name"]
                    
                #download file
                r = requests.get(url, headers={'Authorization': 'Bearer %s' % BOT_TOKEN})
                r.raise_for_status
                file_data = r.content 

                # write and open file in new card doc, check for spreadsheet hits
                with open("new-card.docx" , 'w+b') as f:
                    f.write(bytearray(file_data))
                    check_tup = get_author_and_title("new-card.docx")
                    spreadsheet_search = check_spreadsheet_title(check_tup[0])
                    if spreadsheet_search.empty:
                        say("Please ensure this article is properly claimed in the reserach tracking spreadsheet.")
                    elif len(spreadsheet_search) > 1:
                        say("This article has already been claimed or cut, please DELETE the message with the file in it.")
                    elif spreadsheet_search["Claimed By"].str.contains(first_name).sum() == 0:
                            say("This article has already been claimed or cut, please DELETE the message with the file in it.")
        if docx_counter < 1:
            say("There were no cards in this message")


if __name__ == "__main__":
    SocketModeHandler(app, APP_TOKEN).start()
