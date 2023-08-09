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
    Checks spreadsheet for the article title, if it is in there, returns 1, if not, returns 0. 
    """
    if sheet_data.iloc[:, 3].str.contains(word, case=True).sum() > 0:
        return 1
    else:               
        return 0

def check_spreadsheet_author(word):
    """
    Checks spreadsheet for article author, if it is in there, returns 1, if not, returns 0. 
    """
    if sheet_data.iloc[:, 0].str.contains(word, case=True).sum() > 0:
        return 1
    else:
        return 0

@app.event({"type" : "message", "subtype" : "file_share"})
def get_file(event, say):
        url = event["files"][0]["url_private"]
        file_name = event["files"][0]["name"]

        # download file
        r = requests.get(url, headers={'Authorization': 'Bearer %s' % BOT_TOKEN})
        r.raise_for_status
        file_data = r.content   # get binary content

        # save file to disk
        with open("new-card.docx" , 'w+b') as f:
                f.write(bytearray(file_data))
                check_tup = get_author_and_title("new-card.docx")
                if check_spreadsheet_title(check_tup[0]) + check_spreadsheet_author(check_tup[1]) > 1:
                     say("This card has already been cut, please DELETE the message with the file in it")
                


if __name__ == "__main__":
    SocketModeHandler(app, APP_TOKEN).start()
