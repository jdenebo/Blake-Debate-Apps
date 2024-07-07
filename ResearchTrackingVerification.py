import requests
import docx2txt
import pygsheets
import pandas as pd
from slack_bolt import App
from slack_bolt.adapter.socket_mode import SocketModeHandler
from rapidfuzz import fuzz, utils
from docxcompose.composer import Composer
from docx import Document as Document_compose


APP_TOKEN = APP_TOKEN
#insert your respective tokens
BOT_TOKEN = BOT_TOKEN


app = App(token=BOT_TOKEN)


sheet_id = #whatever sheet ID you want
sheet_name = #whatever sheet you want
url = #your spreadsheet url
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
    perm_rev_str = cite_text[::-1]
    rev_str = cite_text[::-1]
    quotes = []
    starting_idx = 0
    while (chr(8220) in rev_str) or (chr(8221) in rev_str) or (chr(34) in rev_str):
        first = perm_rev_str.find(chr(8220), starting_idx+1)
        second = perm_rev_str.find(chr(8221), starting_idx+1)
        third = perm_rev_str.find(chr(34), starting_idx+1)
        if first > 0:
            quotes.append(first)
        if second > 0:
            quotes.append(second)
        if third > 0:
            quotes.append(third)
        rev_str = rev_str[max(first, second, third)+1:]
        starting_idx += max(first, second, third)
    ordered_quotes = sorted(quotes)
    title = cite_text[len(cite_text)-ordered_quotes[1]:len(cite_text)-ordered_quotes[0]-1]

    
    #need to find initials, so look for ")JDE" or just JDE in the cite
    rev_txt = cite_text[::-1]
    parens_idx = rev_txt.index(")")
    slash_idx = rev_txt.index("/")
    rev_initials = None
    if parens_idx > slash_idx > 0:
        rev_initials = rev_txt[:slash_idx]
    else:
        rev_initials = rev_txt[:parens_idx]
    initials = rev_initials[::-1]

    #need to count number of occurences  of the initials
    search_str_rev = rev_txt[:parens_idx+1]
    search_str = search_str_rev[::-1]
    count = text.count(search_str)


    return (title, initials.strip(), count)

def check_spreadsheet_title(word, sheet_num):
    """
    Checks spreadsheet for the article title, if it is in there, it returns the row that contains it. If its is not, it returns an empty dataframe.
    """
    if sheet_num == 1:
        sheet_data = pd.read_csv(URL_1)
    else:
        sheet_data = pd.read_csv(URL_2)
   # filtered_df = sheet_data[sheet_data.iloc[:, 3].str.contains(word, case=False).fillna(False)]
    sheet_data["matching"] = sheet_data.iloc[:, 2].map(lambda x: fuzz.ratio(str(x), word, processor=utils.default_process))
    # print(sheet_data[sheet_data["matching"] > 90])
    return sheet_data[sheet_data["matching"] > 90]

def check_spreadsheet_author(word, sheet_num):
    """
    Checks spreadsheet for article author, if it is in there, it returns the row that contains it. If its is not, it returns an empty dataframe.
    """
    if sheet_num == 1:
        sheet_data = pd.read_csv(URL_1)
    else:
        sheet_data = pd.read_csv(URL_2)
    filtered_df = sheet_data[sheet_data.iloc[:, 0].str.contains(word, case=False).fillna(False)]
    if filtered_df.empty:
        return 0
    else:
        return filtered_df


def compute_count(topic, initials, card_num):
    topic_cols = {'SEPTOBER': 3,'NOCEMBER': 5, 'JANUARY':7, 'FEBRUARY': 9, 'APRIL' : 11, 'JUNE': 13}
    gc = pygsheets.authorize(service_file='creds.json')
    sh = gc.open('Card Counts 2024-25')
    wks = sh.sheet1
    drange = pygsheets.datarange.DataRange(start='B2', end='B42', worksheet=wks)
    for cell in drange:
        cell_obj = cell[0]
        if cell_obj.row == 41:
            break
        if initials.strip() == cell_obj.value:
            row = cell_obj.row
            article_col = topic_cols[topic]
            card_col = article_col+1
            # prev_val_art = wks.cell((row, article_col)).value
            prev_val_card = wks.cell((row, card_col)).value
            # wks.cell((row,article_col)).set_value(int(prev_val_art) + 1)
            wks.cell((row,card_col)).set_value(int(prev_val_card) + card_num)
            break


@app.event({"type" : "message", "subtype" : "file_share"})
def get_file(event, say, client):
    if "thread_ts" not in event.keys():
        description = event["text"]
        topic = description.split("-")[0].upper().strip()
        topic_sheet = 1
        if SHEET_NAME_2.upper().startswith(topic):
            topic_sheet = 2
        username = client.users_profile_get(user=event["user"])
        first_name = username["profile"]["real_name"].split(" ")[0]
        docx_counter = 0
        for file in event["files"]:
            if file["filetype"] == "docx":
                docx_counter += 1
                url = file["url_private"]
                    
                #download file
                r = requests.get(url, stream=True, headers={'Authorization': 'Bearer ' + BOT_TOKEN})
                r.raise_for_status()
                file_data = r.content

                
                # write and open file in new card doc, check for spreadsheet hits
                with open("new-card.docx" , 'wb') as f:
                    f.write(file_data)
                    check_tup = get_author_and_title("new-card.docx")
                    compute_count(topic, check_tup[1], check_tup[2])
                    spreadsheet_search = check_spreadsheet_title(check_tup[0], topic_sheet)
                    if spreadsheet_search.empty:
                        say("Please ensure this article is properly claimed in the research tracking spreadsheet.")
                    elif len(spreadsheet_search) > 1:
                        say("This article has already been claimed or cut, please DELETE the message with the file in it.")
                    elif spreadsheet_search["Claimed By"].str.contains(first_name, case=False).sum() == 0:
                            say("This article has already been claimed or cut, please DELETE the message with the file in it.")
                            
                composer = Composer(Document_compose("card-comp.docx"))
                composer.append(Document_compose("new-card.docx"))
                composer.save("card-comp.docx")
                

        if docx_counter < 1:
            say("There were no cards in this message")


if __name__ == "__main__":
    SocketModeHandler(app, APP_TOKEN).start()
