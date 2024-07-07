// EVIDENCE COLLECT BOT

// enter your ID
var folder = DriveApp.getFolderById();

// enter your slack token
var slackToken = "";

// enter your webhook
var webhook = "";

function doPost (e) {
  const data = e.postData.contents
  const payload = JSON.parse(data);
  if (payload.type == "url_verification") {
    return ack(payload.challenge) }
  const time = payload.event.ts
  const props = PropertiesService.getUserProperties()
  if (props.getProperty(time) === null) {
    ScriptApp.newTrigger('handlepost').timeBased().after(100).create();
    props.setProperty(time, data);
  }
  return ContentService.createTextOutput("HTTP 200 OK")
}

function doGet(e) {
  return ContentService.createTextOutput("HTTP 200 OK");
}

function handlepost() {
  var triggers = ScriptApp.getProjectTriggers();
  ScriptApp.deleteTrigger(triggers[triggers.length-1])
  const props = PropertiesService.getUserProperties()
  if (props.getKeys().length > 0) {
    var pairs = props.getProperties()
    for (var key in pairs) {
      var temp = pairs[key]
      props.deleteProperty(key)
      processPost(temp)
    }
  }

  if (triggers.length > 10) {
    for (i=2; i<10; i++)
    ScriptApp.deleteTrigger(triggers[i])
  }
}

function processPost(data) {
  var payload = JSON.parse(data)
  if (payload.event.type === "message") {
        if (payload.event.subtype === "file_share" && !("message" in payload.event)) {
            var files = payload.event.files
            var len = files.length
            var split_desc = payload.event.text.split("-");
            for (i=0; i<len; i++) {
              let f = files[i];
              let fileId = f.id;
              let privURL = f.url_private;
              var options = {
                  headers: {
                    'Authorization': 'Bearer ' + slackToken
                    }
                };
              var fileResponse = UrlFetchApp.fetch(privURL, options);
              var blob = fileResponse.getBlob();
              var file = folder.createFile(blob);

                  
              if (f.filetype == "docx") {
                  var gfile = convertToGoogleDoc(file.getBlob());
                  var sheetId = "19cAWZUoMJmtrefutXp9eAjoxDGpk7CcpkrGov_o91_I";
                  var sheetName1 = "Septober2024";
                  var sheetName2 = "June2024";

                  var data1 = fetchSheetData(sheetId, sheetName1);
                  var data2 = fetchSheetData(sheetId, sheetName2);

                  // enter doc ID
                  var docId = "";

                  // Going to try to set up specific folder collection for each card and then check those specific files for title
                  var authorAndTitle = getAuthorAndTitle(gfile);
                  DriveApp.getFileById(gfile).setTrashed(true);

                  var topic = split_desc[0].toUpperCase(); // Example topic
                  var topicSheet = 1;
                  if (sheetName2.toUpperCase().startsWith(topic)) {
                    topicSheet = 2;
                  }

                  var spreadsheetData = (topicSheet === 1) ? data1 : data2;
                  var spreadsheetSearch = checkSpreadsheetTitle(authorAndTitle.title, spreadsheetData);
                  console.log(authorAndTitle.title)
                  console.log(spreadsheetSearch)
                  if (spreadsheetSearch.length === 0) {
                    var message = {"text" : "Please ensure this article is properly claimed in the research tracking spreadsheet.", "thread_ts" : payload.event.ts};
                    slackPost(message);
                  } else if (spreadsheetSearch.length > 1) {
                    var message = {"text" : "This article has already been claimed or cut, please DELETE the message with the file in it.", "thread_ts" : payload.event.ts};
                    slackPost(message);
                    continue;
                  } else {
                    var resp = callWebApi(slackToken, "users.profile.get", `?user=${payload.event.user}`);
                    var firstNameJson = JSON.parse(resp.getContentText());
                    var firstName = firstNameJson.profile.real_name.split(' ')[0]
                    var isClaimedByUser = spreadsheetSearch.some(row => row[5].toLowerCase().includes(firstName.toLowerCase()));
                    if (!isClaimedByUser) {
                      var message = {"text" : "This article has already been claimed or cut, please DELETE the message with the file in it.", "thread_ts" : payload.event.ts};
                      slackPost(message);
                      continue;
                    }
                  }
                  
                  computeCount(topic, authorAndTitle.initials, authorAndTitle.count);
                }
            }
        } else if (payload.event.subtype === "message_deleted") {
          console.log("NEED TO IMPLEMENT MSG DELETION TO SUBTRACT ARTICLE COUNTS")
        }
      }
}


function convertToGoogleDoc(blob, name) {
  var resource = {
    title: name,
    mimeType: MimeType.GOOGLE_DOCS
  };

  var options = {
    convert: true
  };
  
  // Insert the file and convert it to Google Docs format
  var file = Drive.Files.create(resource, blob, options);
  
  return file.id;
}

function fetchSheetData(sheetId, sheetName) {
  var url = `https://docs.google.com/spreadsheets/d/${sheetId}/gviz/tq?tqx=out:csv&sheet=${sheetName}`;
  var response = UrlFetchApp.fetch(url);
  var csvContent = response.getContentText();
  var data = Utilities.parseCsv(csvContent);
  return data;
}

function getAuthorAndTitle(docId) {
  var file = DriveApp.getFileById(docId);
  var doc = DocumentApp.openById(docId);
  var body = doc.getBody().getText();
  var lines = body.split("\n");
  
  var citeText = '';
  for (var i = 0; i < lines.length; i++) {
    if (lines[i].includes("(") && (lines[i].includes('“') || lines[i].includes('"')) && lines[i].includes("http")) {
      citeText = lines[i];
      break;
    }
  }

  // Finding title from cite text
  var reversedCiteText = citeText.split('').reverse().join('');
  var quotes = [];
  var startingIdx = 0;
  var prev=0;
  while (reversedCiteText.includes('“') || reversedCiteText.includes('”') || reversedCiteText.includes('"')) {
    var first = reversedCiteText.indexOf('“', startingIdx + 1);
    var second = reversedCiteText.indexOf('”', startingIdx + 1);
    var third = reversedCiteText.indexOf('"', startingIdx + 1);
    var arr = [first,second,third]
    arr.sort();
    let max=-1
    for (i=0; i<3; i++) {
      if (arr[i] < 0) {
        continue
      }
      let idx = arr[i]
      quotes.push(prev+idx)
      if (idx>max) {
        max = idx
      }
    }
    if (max > 0) {
      prev = max+1
    }
    startingIdx = max+1
    reversedCiteText = reversedCiteText.slice(startingIdx)
  }
  quotes.sort(function(a, b) { return a - b; });
  var title = citeText.substring(citeText.length - quotes[1], citeText.length - quotes[0] - 1);

  // Finding initials
  var reversedText = citeText.split('').reverse().join('');
  var parensIdx = reversedText.indexOf(')');
  var slashIdx = reversedText.indexOf('/');
  var initialsReversed = '';
  
  if (parensIdx > slashIdx && slashIdx > 0) {
    initialsReversed = reversedText.substring(0, slashIdx);
  } else {
    initialsReversed = reversedText.substring(0, parensIdx);
  }
  var initials = initialsReversed.split('').reverse().join('').trim();
  // Counting occurrences
  var searchStrReversed = reversedText.substring(0, parensIdx + 1);
  var searchStr = searchStrReversed.split('').reverse().join('');
  searchStr = escapeRegExp(searchStr);
  var count = (body.match(new RegExp(searchStr, 'g')) || []).length;

  return { title: title, initials: initials, count: count };
}

function escapeRegExp(string) {
  return string.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'); // $& means the whole matched string
}

function slackPost(params) {
  const response = UrlFetchApp.fetch(webhook,
    {
      'method' : 'post',
      'contentType': 'application/json',
      'payload' : JSON.stringify(params),
    }
  );
  return response;
}

function checkSpreadsheetTitle(word, sheetData) {
  var matchingRows = [];
  for (var i = 1; i < sheetData.length; i++) {
    var ratio = getSimilarityRatio(sheetData[i][2], word);
    if (ratio > .9) {
      matchingRows.push(sheetData[i]);
    }
  }
  return matchingRows;
}


function getSimilarityRatio(a, b) {
  var longer = a;
  var shorter = b;
  if (a.length < b.length) {
    longer = b;
    shorter = a;
  }
  var longerLength = longer.length;
  if (longerLength == 0) {
    return 1.0;
  }
  return (longerLength - editDistance(longer, shorter)) / parseFloat(longerLength);
}

function editDistance(s1, s2) {
  s1 = s1.toLowerCase();
  s2 = s2.toLowerCase();
  var costs = new Array();
  for (var i = 0; i <= s1.length; i++) {
    var lastValue = i;
    for (var j = 0; j <= s2.length; j++) {
      if (i == 0)
        costs[j] = j;
      else {
        if (j > 0) {
          var newValue = costs[j - 1];
          if (s1.charAt(i - 1) != s2.charAt(j - 1))
            newValue = Math.min(Math.min(newValue, lastValue),
              costs[j]) + 1;
          costs[j - 1] = lastValue;
          lastValue = newValue;
        }
      }
    }
    if (i > 0)
      costs[s2.length] = lastValue;
  }
  return costs[s2.length];
}

function computeCount(topic, initials, cardNum) {
  var topic = topic.trim();
  var sheetId = "";
  var sheet = SpreadsheetApp.openById(sheetId).getSheets()[0];
  var data = sheet.getDataRange().getValues();
  var topicCols = {'SEPTOBER': 3, 'NOCEMBER': 5, 'JANUARY': 7, 'FEBRUARY': 9, 'APRIL': 11, 'JUNE': 13};

  for (var i = 1; i < data.length; i++) {
    if (data[i][1].trim() == initials.trim()) {
      var articleCol = topicCols[topic.toUpperCase()];
      var cardCol = articleCol + 1;
      var prevValCard = parseInt(data[i][cardCol-1]);
      console.log(cardCol, i)
      sheet.getRange(i + 1, cardCol).setValue(prevValCard + cardNum);
      break;
    }
  }
}

function callWebApi(token, apiMethod, params) {
  const response = UrlFetchApp.fetch(
    `https://www.slack.com/api/${apiMethod}${params}`,
    {
      'method' : 'get',
      'contentType': 'application/x-www-form-urlencoded',
      'headers' : { "Authorization": `Bearer ${token}` },
    }
  );
  return response;
}


function ack(payload) {
  const reply = "HTTP 200 OK\nContent-type: text/plain\n" + payload
  return ContentService.createTextOutput(reply);
}
