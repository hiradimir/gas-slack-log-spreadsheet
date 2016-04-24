/**** Do not edit below unless you know what you are doing ****/

const COL_LOG_TIMESTAMP = 1;
const COL_LOG_USER = 2;
const COL_LOG_TEXT = 3;
const COL_LOG_RAW_JSON = 4;
const COL_MAX = COL_LOG_RAW_JSON;

const FETCH_STATUS_START = "START";
const FETCH_STATUS_ARCHIVED = "ARCHIVED";
const FETCH_STATUS_END = "END";

// Slack offers 10,000 history logs for free plan teams
const MAX_HISTORY_PAGINATION = 10;
const HISTORY_COUNT_PER_PAGE = 1000;
// 4分をリミットとする
const TRIGGER_LIMIT = 4 * (60 * 1000);

// Configuration: Obtain Slack web API token at https://api.slack.com/web
const API_TOKEN = PropertiesService.getScriptProperties().getProperty('slack_api_token');
if (!API_TOKEN) {
  throw 'You should set "slack_api_token" property from [File] > [Project properties] > [Script properties]';
}
const APP_WORKSHEET_ID = PropertiesService.getScriptProperties().getProperty('app_worksheet_id');
if (!APP_WORKSHEET_ID) {
  throw 'You should set "app_worksheet_id" property from [File] > [Project properties] > [Script properties]';
}
const LOG_FOLDER_ID = PropertiesService.getScriptProperties().getProperty('log_folder_id');
if (!LOG_FOLDER_ID) {
  throw 'You should set "log_folder_id" property from [File] > [Project properties] > [Script properties]';
}

const LOG_LEVELS = ["trace", "debug", "info", "warn", "error", "fatal"];

var LOG_LEVEL:number = LOG_LEVELS.indexOf("info");
var LOG_LEVEL_PROP = PropertiesService.getScriptProperties().getProperty('log_level');
if (LOG_LEVELS.indexOf(LOG_LEVEL_PROP) >= 0) {
  LOG_LEVEL = LOG_LEVELS.indexOf(LOG_LEVEL_PROP);
}

class SpreadsheetLogger {
  constructor(public id: string) {
  }

  sh: GoogleAppsScript.Spreadsheet.Sheet;

  log_sheet_() {
    var sheet_name = 'log';
    // var ss = SpreadsheetApp.getActiveSpreadsheet();
    var ss = SpreadsheetApp.openById(this.id);

    if (this.sh == null) {
      var sh = ss.getSheetByName(sheet_name);
      if (sh == null) {
        var sheet_num = ss.getSheets().length;
        sh = ss.insertSheet(sheet_name, sheet_num);
        sh.getRange('A1:C1').setValues([['timestamp', 'level', 'message']]).setBackground('#cfe2f3').setFontWeight('bold');
        sh.getRange('A2:C2').setValues([[new Date(), 'info', sheet_name + ' has been created.']]).clearFormat();
      }
      this.sh = sh;
    }
    return this.sh;
  }

  private log_(level: string, message: string) {
    if(LOG_LEVELS.indexOf(level) < LOG_LEVEL) {
      return;
    }

    var sh = this.log_sheet_();
    var now = new Date();
    var last_row = sh.getLastRow();
    sh.insertRowAfter(last_row).getRange(last_row + 1, 1, 1, 3).setValues([[now, level, "'" + message]]);
    return sh;
  }

  trace(message: string) {
    this.log_('trace', message);
  }

  debug(message: string) {
    this.log_('debug', message);
  }

  info(message: string) {
    this.log_('info', message);
  }

  warn(message: string) {
    this.log_('warn', message);
  }

  error(message: string) {
    this.log_('error', message);
  }

  fatal(message: string) {
    this.log_('fatal', message);
  }

}

const myLogger: SpreadsheetLogger = new SpreadsheetLogger(APP_WORKSHEET_ID);

class ChannelLogStatus {
  constructor(public key: string, public timestamp: Date, public status: string) {
  }

  toObjectArray(): Object[] {
    return [this.key, this.timestamp, this.status]
  }
}

class SpreadsheetKeyValueStore {

  constructor(public id: string) {
    this.init();
  }

  sh: GoogleAppsScript.Spreadsheet.Sheet;

  keyStatusMap: {[key: string]: {index: number, values?: ChannelLogStatus}} = {};

  private init() {
    var sh = this.getSheet();
    var values: Object[][] = sh.getSheetValues(1, 1, sh.getMaxRows() - 1, sh.getMaxColumns());
    values.forEach((v, i) => {
      var status = new ChannelLogStatus(<string>v[0], <Date>v[1], <string>v[2]);
      this.keyStatusMap[status.key] = {index: i + 1, values: status};
    });
    return sh;
  }

  private getSheet() {

    if (this.sh == null) {
      var sheet_name = 'keyValue';
      var ss = SpreadsheetApp.openById(this.id);
      var sh = ss.getSheetByName(sheet_name);
      if (sh == null) {
        var sheet_num = ss.getSheets().length;
        sh = ss.insertSheet(sheet_name, sheet_num);
        sh.getRange('A1:C1').setValues([['key', 'timestamp', 'value']]).setBackground('#cfe2f3').setFontWeight('bold');
      }
      this.sh = sh;
    }
    return this.sh;
  }

  private newRow(key: string): number {
    var sh = this.getSheet();
    var last_row = sh.getLastRow();
    sh.insertRowAfter(last_row);
    this.keyStatusMap[key] = {index: last_row + 1, values: null};
    return last_row + 1;
  }

  setStatus(key: string, status: string) {
    if (!this.keyStatusMap[key]) {
      this.newRow(key);
    }
    var keyInfo = this.keyStatusMap[key];
    var sh = this.getSheet();

    var now = new Date();
    keyInfo.values = new ChannelLogStatus(key, now, status);
    sh.getRange(keyInfo.index, 1, 1, 3).setValues([keyInfo.values.toObjectArray()]).clearFormat();
  }

  getStatus(key: string): ChannelLogStatus {
    if (this.keyStatusMap[key]) {
      return this.keyStatusMap[key].values;
    }
    return new ChannelLogStatus(key, new Date(1), "");
  }
}

const keyValueStore = new SpreadsheetKeyValueStore(APP_WORKSHEET_ID);


interface ISlackResponse {
  ok: boolean;
  error?: string;
}

// https://api.slack.com/methods/channels.list
interface ISlackChannelsListResponse extends ISlackResponse {
  channels: ISlackChannel[];
}

// https://api.slack.com/methods/channels.history
interface ISlackChannelsHistoryResponse extends ISlackResponse {
  latest?: string;
  oldest?: string;
  has_more: boolean;
  messages: ISlackMessage[];
}

// https://api.slack.com/methods/users.list
interface ISlackUsersListResponse extends ISlackResponse {
  members: ISlackUser[];
}

// https://api.slack.com/types/channel
interface ISlackChannel {
  id: string;
  name: string;
  created: number;
  is_archived: boolean;
  is_channel: boolean;
  is_general: boolean;

  // ...and more fields
}

// https://api.slack.com/events/message
interface ISlackMessage {
  type: string;
  ts: string;
  user: string;
  text: string;

  // https://api.slack.com/events/message/bot_message
  username?: string;

  // ...and more fields
}

// https://api.slack.com/types/user
interface ISlackUser {
  id: string;
  name: string;

  // ...and more fields
}

// https://api.slack.com/methods/team.info
interface ISlackTeamInfoResponse extends ISlackResponse {
  team: {
    id: string;
    name: string;
    domain: string;
    // ...and more fields
  };
}

function StoreLogsDelta() {
  let logger = new SlackChannelHistoryLogger();
  myLogger.info("Start logger.run");
  logger.run();
  myLogger.info("End   logger.run");
}

interface ISpreadsheetInfo {
  spreadsheet: GoogleAppsScript.Spreadsheet.Spreadsheet;
  sheets: { [ id: string ]: GoogleAppsScript.Spreadsheet.Sheet; };
}
;


class SlackChannelHistoryLogger {
  memberNames: { [id: string]: string } = {};
  teamName: string;

  cachedSpreadSheet: { [id: string]: GoogleAppsScript.Spreadsheet.Spreadsheet } = {};
  cachedSheet: { [id: string]: {[id: string]: GoogleAppsScript.Spreadsheet.Sheet} } = {};

  constructor() {
  }

  requestSlackAPI(path: string, params: { [key: string]: any } = {}): ISlackResponse {
    let url = `https://slack.com/api/${path}?`;
    let qparams = [`token=${encodeURIComponent(API_TOKEN)}`];
    for (let k in params) {
      qparams.push(`${encodeURIComponent(k)}=${encodeURIComponent(params[k])}`);
    }
    url += qparams.join('&');

    myLogger.debug(`==> GET ${url}`);

    let resp = UrlFetchApp.fetch(url);
    let data = <ISlackResponse>JSON.parse(resp.getContentText());
    if (data.error) {
      throw `GET ${path}: ${data.error}`;
    }

    myLogger.debug(`<== GOT`);
    return data;
  }

  run() {
    //時刻格納用の変数
    var starttime = +new Date();

    let usersResp = <ISlackUsersListResponse>this.requestSlackAPI('users.list');
    usersResp.members.forEach((member) => {
      this.memberNames[member.id] = member.name;
    });

    let teamInfoResp = <ISlackTeamInfoResponse>this.requestSlackAPI('team.info');
    this.teamName = teamInfoResp.team.name;

    let channelsResp = <ISlackChannelsListResponse>this.requestSlackAPI('channels.list');

    const channelFetchTime = (ch: ISlackChannel) => {
      var sheetName = this.sheetName(ch);
      var status = keyValueStore.getStatus(sheetName);
      var time = status.timestamp.getTime();
      if (status.status != FETCH_STATUS_END) {
        time = 0;
      }
      return time;
    };

    channelsResp.channels.sort((ch1, ch2)=> {
      var time1 = channelFetchTime(ch1);
      var time2 = channelFetchTime(ch2);
      return time1 - time2;
    });

    for (let ch of channelsResp.channels) {
      this.importChannelHistoryDelta(ch);
      var endtime = +new Date();
      if (endtime - starttime > TRIGGER_LIMIT) {
        myLogger.warn(`TERMINATE by limit time ${endtime - starttime} > ${TRIGGER_LIMIT}`);
        break;
      }

    }
  }

  cachedFolder: GoogleAppsScript.Drive.Folder;

  getLogsFolder(): GoogleAppsScript.Drive.Folder {

    if (!this.cachedFolder) {
      let folder = DriveApp.getFolderById(LOG_FOLDER_ID);
      if (!folder) {
        throw 'You should set "log_folder_id" property from [File] > [Project properties] > [Script properties]';
      }
      this.cachedFolder = folder;

    }
    return this.cachedFolder;
  }

  getSpreadSheet(ch: ISlackChannel, d: Date|string, readonly: boolean = false): GoogleAppsScript.Spreadsheet.Spreadsheet {
    let dateString: string;
    if (d instanceof Date) {
      dateString = this.formatDate(d);
    } else {
      dateString = '' + d;
    }

    let spreadsheet: GoogleAppsScript.Spreadsheet.Spreadsheet;

    let spreadsheetName = dateString;

    if (this.cachedSpreadSheet[spreadsheetName]) {
      spreadsheet = this.cachedSpreadSheet[spreadsheetName];
    } else {
      let folder = this.getLogsFolder();
      let it = folder.getFilesByName(spreadsheetName);
      if (it.hasNext()) {
        let file = it.next();
        spreadsheet = SpreadsheetApp.openById(file.getId());
      } else {
        if (readonly) return null;

        spreadsheet = SpreadsheetApp.create(spreadsheetName);
        folder.addFile(DriveApp.getFileById(spreadsheet.getId()));
      }
    }

    return spreadsheet;
  }

  sheetName(ch: ISlackChannel): string {
    let sheetName = `${ch.name} (${ch.id})`;
    return sheetName;
  }

  getSheet(ch: ISlackChannel, d: Date|string, readonly: boolean = false): GoogleAppsScript.Spreadsheet.Sheet {
    let spreadsheet: GoogleAppsScript.Spreadsheet.Spreadsheet = this.getSpreadSheet(ch, d, readonly);

    if (!spreadsheet) {
      return null;
    }

    let dateString: string;
    if (d instanceof Date) {
      dateString = this.formatDate(d);
    } else {
      dateString = '' + d;
    }

    let sheetByID: { [id: string]: GoogleAppsScript.Spreadsheet.Sheet };

    if (this.cachedSheet[dateString]) {
      sheetByID = this.cachedSheet[dateString];
    } else {
      sheetByID = {};

      let sheets = spreadsheet.getSheets();
      sheets.forEach((s: GoogleAppsScript.Spreadsheet.Sheet) => {
        let name = s.getName();
        let m = /^(.+) \((.+)\)$/.exec(name); // eg. "general (C123456)"
        if (!m) return;
        sheetByID[m[2]] = s;
      });
      this.cachedSheet[dateString] = sheetByID;
    }

    let sheet = sheetByID[ch.id];
    if (!sheet) {
      if (readonly) return null;
      sheet = spreadsheet.insertSheet();
      sheet.setColumnWidth(COL_LOG_TEXT, 800);
    }

    let sheetName = this.sheetName(ch);
    if (sheet.getName() !== sheetName) {
      sheet.setName(sheetName);
    }

    return sheet;
  }

  importChannelHistoryDelta(ch: ISlackChannel) {
    myLogger.info(`importChannelHistoryDelta ${ch.name} (${ch.id})`);
    let sheetName = this.sheetName(ch);
    var prevStatus = keyValueStore.getStatus(sheetName);

    if (prevStatus.status == FETCH_STATUS_ARCHIVED) {
      return;
    }

    keyValueStore.setStatus(sheetName, FETCH_STATUS_START);

    let now = new Date();
    let oldest = '1'; // oldest=0 does not work

    let existingSheet = this.getSheet(ch, now, true);
    if (!existingSheet) {
      // try previous month
      now.setMonth(now.getMonth() - 1);
      existingSheet = this.getSheet(ch, now, true);
    }
    if (existingSheet) {
      let lastRow = existingSheet.getLastRow();
      try {
        let data = <ISlackMessage>JSON.parse(<string>existingSheet.getRange(lastRow, COL_LOG_RAW_JSON).getValue());
        oldest = data.ts;
      } catch (e) {
        myLogger.warn(`while trying to parse the latest history item from existing sheet: ${e}`)
      }
    }

    let messages = this.loadMessagesBulk(ch, {oldest: oldest});
    let dateStringToMessages: { [dateString: string]: ISlackMessage[] } = {};

    messages.forEach((msg) => {
      let date = new Date(+msg.ts * 1000);
      let dateString = this.formatDate(date);
      if (!dateStringToMessages[dateString]) {
        dateStringToMessages[dateString] = [];
      }
      dateStringToMessages[dateString].push(msg);
    });

    for (let dateString in dateStringToMessages) {
      let sheet = this.getSheet(ch, dateString);

      var timezone = sheet.getParent().getSpreadsheetTimeZone();
      var lastTS: number = 0;
      let lastRow = sheet.getLastRow();
      if (lastRow > 0) {
        try {
          let data = <ISlackMessage>JSON.parse(<string>sheet.getRange(lastRow, COL_LOG_RAW_JSON).getValue());
          lastTS = +data.ts || 0;
        } catch (_) {
        }
      }

      let rows = dateStringToMessages[dateString].filter((msg) => {
        return +msg.ts > lastTS;
      }).map((msg) => {
        let date = new Date(+msg.ts * 1000);
        return [
          Utilities.formatDate(date, timezone, 'yyyy-MM-dd HH:mm:ss'),
          this.memberNames[msg.user] || msg.username,
          this.unescapeMessageText(msg.text),
          JSON.stringify(msg)
        ]
      });
      if (rows.length > 0) {
        let range = sheet.insertRowsAfter(lastRow || 1, rows.length)
          .getRange(lastRow + 1, 1, rows.length, COL_MAX);
        range.setValues(rows);
      }
    }
    if (ch.is_archived) {
      keyValueStore.setStatus(sheetName, FETCH_STATUS_ARCHIVED);
    } else {
      keyValueStore.setStatus(sheetName, FETCH_STATUS_END);
    }

  }

  formatDate(dt: Date): string {
    return Utilities.formatDate(dt, Session.getScriptTimeZone(), 'yyyy-MM');
  }

  loadMessagesBulk(ch: ISlackChannel, options: { [key: string]: string|number } = {}): ISlackMessage[] {
    let messages: ISlackMessage[] = [];

    // channels.history will return the history from the latest to the oldest.
    // If the result's "has_more" is true, the channel has more older history.
    // In this case, use the result's "latest" value to the channel.history API parameters
    // to obtain the older page, and so on.
    options['count'] = HISTORY_COUNT_PER_PAGE;
    options['channel'] = ch.id;
    let loadSince = (oldest?: string) => {
      if (oldest) {
        options['oldest'] = oldest;
      }
      // order: recent-to-older
      let resp = <ISlackChannelsHistoryResponse>this.requestSlackAPI('channels.history', options);
      messages = resp.messages.concat(messages);
      return resp;
    }

    let resp = loadSince();
    let page = 1;
    while (resp.has_more && page <= MAX_HISTORY_PAGINATION) {
      myLogger.info(`channels.history.pagination ${ch.name} (${ch.id}) ${page}`);
      resp = loadSince(resp.messages[0].ts);
      page++;
    }

    // oldest-to-recent
    return messages.reverse();
  }

  unescapeMessageText(text?: string): string {
    return (text || '')
      .replace(/&lt;/g, '<')
      .replace(/&gt;/g, '>')
      .replace(/&quot;/g, '"')
      .replace(/&amp;/g, '&')
      .replace(/<@(.+?)>/g, ($0, userID) => {
        let name = this.memberNames[userID];
        return name ? `@${name}` : $0;
      });
  }
}
