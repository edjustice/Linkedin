const fs = require('fs');
const prompt = require('prompt');
const Nightmare = require('nightmare');
const xlsx = require('xlsx');

let output_data_dir = "stored_data/";

let connections = [];
let extractedData = {
  extracted_data: []
};
let nightmare;

function sleep(ms){
    return new Promise(resolve=>{
        setTimeout(resolve,ms)
    })
}

// 30 extracts by session worked fine for exporting 3000 contacts' emails & phones
let max_extracts_by_session = 30;
let subloop_count = 0;
  

// Setup prompt attributes
let prompt_attrs = [
  { 
    name: 'email', 
    required: true, 
    message: "LinkedIn email" 
  },
  { 
    name: 'password',
    hidden: true,
    required: true,
    message: "LinkedIn password"
  },
  {
    name: 'searchInterval',
    default: "100",
    message: "Wait interval between each connection search (in ms)"
  }
]

// Define variables
let user_email, password, showNightmare, searchInterval, getUsersPhone, getUsersSummary, getUsersLocation, getLinkedinProfile;
let emails = [];
let xls;
let index = 0;

// This function starts the process by asking user for LinkedIn credentials, as well config options
// - email & password are used to log in to linkedin
function start() {
  xls = new XLSXReader(xlsx, "Recipients.xlsx", 0);
  if(xls.lines.length == 0)
    return;
  if(process.argv.length >= 4)
  {
    prompt_attrs[0] = prompt_attrs[2];
    prompt_attrs.pop();
    prompt_attrs.pop();
    user_email = process.argv[2];
    password = process.argv[3];
  }
  prompt.start()
  
  prompt.get(prompt_attrs, (err, result) => {
    if(process.argv.length < 4){
      user_email = result.email
      password = result.password
    }
    showNightmare = result.showNightmare === "yes"
    searchInterval = parseInt(result.searchInterval)
    nightmare = Nightmare({
      show: true,
      waitTimeout: 35000
    })
    sendMessages(index);
  })
}


// Emails are stored in this array to be written to email.txt later.
let result = []
let phones = []
let summaries = []
let locations = []
let profiles = []
let sub_result = []
let sub_phones = []
let sub_summaries = []
let sub_locations = []
let sub_profiles = []
let failed = []
let failed_sub = []

let email, phone, summary, location, profile;
let allData = [];
let sub_allData = [];

// Initial email extraction procedure
// Logs in to linked in and runs the getEmail async function to actually extract the emails
async function sendMessages(index,reset=false) {
  sub_result = [];
  sub_phones = [];
  sub_summaries = [];
  sub_locations = [];
  sub_profiles = [];
  sub_allData = [];
  failed_sub = [];
  if(reset){
    await nightmare.end()
    nightmare = Nightmare({
      show: showNightmare,
      waitTimeout: 35000
    })
  }
  try {
    await login()
    await nightmare
    .wait('#messaging-tab-icon')
    .click("#messaging-tab-icon")
    .wait(1000)
    .run(async () => {
      for(var i = 0; i < xls.lines.length; ++i)
      {
        // console.log(xls.lines, i, xls.lines[i])
        await sendMessage(xls.lines[i].firstname + " " + xls.lines[i].lastname, xls.lines[i].message, xls.lines[i].pj.length != 0, i);
        await nightmare
        .wait(200)
        .click(".msg-form__send-button");
        await sleep(searchInterval);
      }
      var end_fn = async () => {
        await nightmare.end();
      }
      xls.writeAsyncCallback = end_fn;
      xls.writeResponsesColumn();
    })
  } catch(e) {
    console.error("An error occured while attempting to login to linkedin.");
  }
}

async function login()
{
  return await nightmare
  .goto('https://linkedin.com')
  .exists("#login-email")
  .then(async function(result){
    if(result)
    {
      return await nightmare
      .insert('#login-email', user_email)
      .insert('#login-password', password)
      .click('#login-submit')
    }
    else
    {
      return await nightmare
      .click(".nav__button-secondary")
      .wait("#username")
      .insert('#username', user_email)
      .insert('#password', password)
      .click('.btn__primary--large')
    }
  })
}

async function sendMessage(peopleName, message, pj=false, i=0, count=0) {


  // Condition is here to make sure no more than the limit of mails is extracted on each interval
  if (count < 10) {
        try {
      return await nightmare
        .wait(2000)
        .click("li-icon[type=compose-icon]")
        .insert('.msg-connections-typeahead__search-field', '')
        .insert('.msg-connections-typeahead__search-field', peopleName)
        .wait(".msg-connections-typeahead__search-field", 2000)
        .click('.msg-connections-typeahead__search-field')// press enter
        .wait(".msg-connections-typeahead__search-results",2000)
        // .exists(".msg-connections-typeahead__search-results > li > button")
        .exists(".msg-connections-typeahead__search-results")
        .then(async function(result){
          if(result)
          {
            console.log(peopleName, "is ok")
            xls.lines[i].response = "ok";
            return await nightmare
              // .wait(1000)
              .click('.msg-connections-typeahead__search-field')
              .type('body', '\u000d')
              .wait(".msg-form__contenteditable",1000)
              .insert(".msg-form__contenteditable", message)
              .type('body', '\u000d')
              .click(".msg-form__send-button")
              .wait(10000)
              // .click("li-icon[type=compose-icon]")
              // .click(".msg-form__send-button")
          }
          else
          {
            console.log(peopleName, "is not found !!!")
            xls.lines[i].response = "user_not_found";
            return;
          }

        })

    } catch(e) {
      console.log(e);
      console.error("Erreur");
      return;
    }

  } else {
    // When all emails have been extracted, end nightmare crawler and add emails to email.txt
    await nightmare.end();
  }
}

class XLSXReader
{
  filename;
  worksheet;
  workbook;
  headers;
  lines;
  writeAsyncCallback;
  neededHeaders = ["firstname", "lastname", "message", "pj"]; // required headers

  constructor(XLSX, filename, sheetID=0)
  {
    this.filename = filename;
    this.XLSX = XLSX;
    try
    {
      this.workbook = XLSX.readFile(filename);
    }catch(e){
      console.error("Cannot open XLS file \"" + this.filename + "\"");
      process.exit()
    }

    try
    {
      this.worksheet = this.workbook.Sheets[this.workbook.SheetNames[0]];
    }catch(e){
      console.error("Cannot open specified sheet in XLS file");
      process.exit()
    }

    this.ReadHeaders();
    this.BuildLines();
    if(!this.CheckHeaders())
      throw new Error("Excel file does not have needed columns.\nThey must be, at least (no-casse sensible) : " + this.neededHeaders.join(", ") + ".\n");
  }

  ReadHeaders()
  {
    var headers = [];
    var range = this.XLSX.utils.decode_range(this.worksheet['!ref']);
    var C, R = range.s.r;
    for(C = range.s.c; C <= range.e.c; ++C) {
        var cell = this.worksheet[this.XLSX.utils.encode_cell({c:C, r:R})] /* find the cell in the first row */
  
        var hdr = "UNKNOWN " + C; // <-- replace with your desired default 
        if(cell && cell.t) hdr = this.XLSX.utils.format_cell(cell);
  
        // headers.push(hdr);
        headers.push(hdr.toLowerCase());
    }
    this.headers = headers;
    return headers;
  }

  BuildLines()
  {
    var lines = [];
    var range = this.XLSX.utils.decode_range(this.worksheet['!ref']);
    var R, C;
    for(R = range.s.r+1; R <= range.e.r; ++R) {
        var obj = {};

        for(C = range.s.c; C <= range.e.c; ++C)
        {
          var cell = this.worksheet[this.XLSX.utils.encode_cell({c:C, r:R})]
          var hdr = "UNKNOWN " + C; 
          if(cell && cell.t) hdr = this.XLSX.utils.format_cell(cell);
          obj[this.headers[C]] = hdr;
        }

        lines.push(obj);
    }
    this.lines = lines;
    return lines;
  }

  CheckHeaders()
  {
    var success = true;
    for(var i = 0; i < this.neededHeaders.length; ++i)
    {
      if(this.headers.indexOf(this.neededHeaders[i]) == -1)
      {
        success = false;
        break;
      }
    }

    return success;
  }

  /* Custom */

  AppendColumnToFile(column)
  {
    var range = this.XLSX.utils.decode_range(this.worksheet['!ref']);
    range.e.c++;
    var new_column_id = range.e.c;
    for(var i = 0; i < column.length; ++i)
    {
      this.worksheet[this.XLSX.utils.encode_cell({c: new_column_id, r: i})] = {t: 's', v: column[i]};
    }

    this.worksheet['!ref'] = this.XLSX.utils.encode_range(range);
    this.XLSX.writeFileAsync(this.filename, this.workbook, this.writeAsyncCallback ? this.writeAsyncCallback : (err, res) => {});
  }

  writeResponsesColumn()
  {
    console.log(this.lines);
    var header = new Date().toISOString();
    var column = [header];
    for(var i = 0; i < this.lines.length; ++i)
    {
      column.push(this.lines[i].response ? this.lines[i].response : "strange_error: field \"response\" is missing in row data");
    }

    this.AppendColumnToFile(column);
  }
}

start();