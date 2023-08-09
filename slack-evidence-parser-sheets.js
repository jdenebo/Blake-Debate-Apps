
// paste your bot token string
var token = TOKEN;

// gets the sheet the code is operating on
var ss = SpreadsheetApp.getActiveSpreadsheet();

// can change to whichever sheet is preferred
var sheet = ss.getSheetByName('Sheet1');

//paste webhook in if you want to use the webhook messaging method
var webhook = HOOK;

function doPost(e) {
  const payload = JSON.parse(e.postData.contents);
  if (typeof(payload.challenge) !== "undefined") {
    return ack(payload.challenge) }
  
  var last_row = sheet.getLastRow();

  if (payload.event.type === "message") {
      if (payload.event.subtype === "file_share") {

        //handling file names, vars for pdf
        var names = [];
        var ids = [];
        var len = payload.event.files.length;
        var pdf = 0;
        for (i=0; i < len; i++) {
          names.push(payload.event.files[i].name);
          ids.push(payload.event.files[i].id);
          if (payload.event.files[i].filetype == "pdf" && (!pdf)) {
            pdf = 1;
          }
        }

        if (!pdf) {
          var message1 = {"text" : "The PDF is missing, if it is available please delete this post and try again!"};
          slackPost(message1);
        } else if (len <= 1) {
          var message1 = {"text" : "The cards are missing, please delete this post and try again!"}
          slackPost(message1);
        } else {
    
        var date = payload.event.ts;
        sheet.getRange(last_row+1, 10).setValue(parseFloat(date));
        var user = payload.event.user;
        sheet.getRange(last_row+1, 11).setValue(user);

        //viewable date set
        sheet.getRange(last_row+1, 2).setValue(date_time(date));



        // fill out ids array
        const id_str = ids.toString();
        sheet.getRange(last_row+1, 12).setValue(id_str);


        if (pdf != 0) {
          sheet.getRange(last_row+1, 3).setValue("YES");
        } else {
          sheet.getRange(last_row+1, 3).setValue("NO");
        }

        // setting file name to just the [johnson 22] format
        sheet.getRange(last_row+1, 1).setValue(names[0].split(".")[0]);


        // description
        var description_exists = payload.event.text.length;
        if (!payload.event.text.length) {
          var message2 = {"text" : "You are missing the description of the file, please EDIT this message to include it in the following format: [Topic(Septober, Nocember, April etc)] - [Year] - [Argument Type(link, impact, etc)] - [Description of Cards]"};
          slackPost(message2);
        } else {
          sheet.getRange(last_row+1, 4).setValue(payload.event.text);
        }
      
        //permalink
        var permalink = callWebApi(token, "chat.getPermalink", `?channel=${payload.event.channel}&message_ts=${date}`);
        var data = JSON.parse(permalink.getContentText());
        sheet.getRange(last_row+1, 6).setValue(data.permalink);


        //user conversions
        var username = callWebApi(token, "users.profile.get", `?user=${user}`);
        var data2 = JSON.parse(username.getContentText());
        sheet.getRange(last_row+1, 5).setValue(data2.profile.real_name);


        //description parse
        var split_desc = payload.event.text.split("-");
        if (split_desc.length < 4 && description_exists) {
          var message3 = {"text" : "Your description is missing certain elements, please EDIT this message to correct it in the following format: [Topic(Septober, Nocember, April etc)] - [Year] - [Argument Type(link, impact, etc)] - [Description of Cards]"};
          slackPost(message3);
        }

        sheet.getRange(last_row+1, 7).setValue(split_desc[0].trim() + " " + split_desc[1].trim());
        sheet.getRange(last_row+1, 8).setValue(split_desc[1].trim());


        // arg type, maybe more work to be done to handle multiple types?
        var type = split_desc[2].trim();
        check_arg_type(type, last_row+1);
        
        }
  





      }  else if (payload.event.subtype === "message_changed") {

      //changing rows based on message updates
      var val = payload.event.message.ts;
      var row_to_update = val_search(val);
      if (row_to_update){
        sheet.getRange(row_to_update, 4).setValue(payload.event.message.text);
      }

      //description parse
      var split_desc = payload.event.message.text.split("-");
      sheet.getRange(row_to_update, 7).setValue(split_desc[0].trim() + " " + split_desc[1].trim());
      sheet.getRange(row_to_update, 8).setValue(split_desc[1].trim());


      // arg type, maybe more work to be done to handle multiple types?
      var type = split_desc[2].trim();
      check_arg_type(type, row_to_update)





    } else if (payload.event.subtype === "message_deleted") {
        var val = payload.event.deleted_ts;
        var row_to_update = val_search(val);
        if (row_to_update) {
          sheet.deleteRows(row_to_update, 1);
      }





    } else if (payload.event.subtype === "message_replied") {
        sheet.getRange("D30").setValue(1000)
        var parent = val_search(payload.event.message.thread_ts);
        var ids_arr= [];
        // for (i=0; i<payload.event.message.reply_count; i++) {
        //     ids_arr.push()
        // }
        sheet.getRange(parent, 13).setValue(payload.event.message.text);

    }
   else if (payload.event.type === "file_deleted") {
    // deletes spreadsheet inputs if files are deleted
    const row_to_delete = val_search(payload.event.file_id);
    if (row_to_delete) {
      sheet.deleteRows(row_to_delete, 1); }
  }
  }
}


function check_arg_type(type, row) {
  if (type == "impact" || type == "Impact" || type == "IMPACT" || type == "i" || type == "I" || type == "IMP" || type == "imp" || type == "Imp") {
    sheet.getRange(row, 9).setValue("Impact");
  } else if (type == "link" || type == "Link" || type == "LINK" || type == "l" || type == "L") {
    sheet.getRange(row, 9).setValue("Link");
  } else if (type == "IL" || type == "il" || type == "internal link" || type == "INTERNAL LINK" || type == "Il" || type == "iL" || type == "Internal Link" || type == "internal Link" || type == "Internal link") {
    sheet.getRange(row, 9).setValue("Internal Link");
  } else if (type == "UQ" || type == "uq" || type == "uQ" || type == "Uq" || type == "Uniqueness" || type == "uniqueness" || type == "UNIQUENESS" || type.includes("unq") || type.includes("Unq") || type.includes("uniq") || type.includes("Uniq")) {
    sheet.getRange(row, 9).setValue("Uniqueness");
  } else if (type == "defense" || type == "Defense" || type == "def" || type == "DEF" || type == "DEFENSE" || type == "D" || type == "d"){
    sheet.getRange(row, 9).setValue("Defense");
  } else if (type == "impact d" || type == "impact D" || type == "impact defense" || type == "Impact Defense" || type == "Impact d" || type == "Impact D") {
    shee.getRange(row, 9).setValue("Impact D");
  } else {
    sheet.getRange(row, 9).setValue(type);
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

function callWebApiPOST(token, apiMethod, params) {
  const response = UrlFetchApp.fetch(
    `https://www.slack.com/api/${apiMethod}`,
    {
      'method' : 'post',
      'contentType': 'application/json',
      'headers' : { "Authorization": `Bearer ${token}` },
      'payload' : JSON.stringify(params),
    }
  );
  return response;
}

function date_time (t) {
  var date = new Date(t*1000);
  var formattedDate = Utilities.formatDate(date, "GMT-6:00", "MM/dd/yyyy");
  return formattedDate;
}


function val_search (v) {
  const dataRange = sheet.getDataRange();
  const values = dataRange.getValues();
  const columnIndex = 9 // INDEX OF COLUMN FOR COMPARISON CELL
  const matchText = parseFloat(v);
  const index = values.findIndex(row => row[columnIndex] === matchText)
  const rowNumber = index + 1
  return rowNumber
}

function ack(payload) {
  if (typeof payload === "string") {
    return ContentService.createTextOutput(payload);
  } else {
    return ContentService.createTextOutput(JSON.stringify(payload))
               .setMimeType(ContentService.MimeType.JSON);
  }
}
