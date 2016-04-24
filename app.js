/**** Do not edit below unless you know what you are doing ****/
var COL_LOG_TIMESTAMP = 1;
var COL_LOG_USER = 2;
var COL_LOG_TEXT = 3;
var COL_LOG_RAW_JSON = 4;
var COL_MAX = COL_LOG_RAW_JSON;
var FETCH_STATUS_START = "START";
var FETCH_STATUS_ARCHIVED = "ARCHIVED";
var FETCH_STATUS_END = "END";
// Slack offers 10,000 history logs for free plan teams
var MAX_HISTORY_PAGINATION = 10;
var HISTORY_COUNT_PER_PAGE = 1000;
// 4分をリミットとする
var TRIGGER_LIMIT = 4 * (60 * 1000);
// Configuration: Obtain Slack web API token at https://api.slack.com/web
var API_TOKEN = PropertiesService.getScriptProperties().getProperty('slack_api_token');
if (!API_TOKEN) {
    throw 'You should set "slack_api_token" property from [File] > [Project properties] > [Script properties]';
}
var APP_WORKSHEET_ID = PropertiesService.getScriptProperties().getProperty('app_worksheet_id');
if (!APP_WORKSHEET_ID) {
    throw 'You should set "app_worksheet_id" property from [File] > [Project properties] > [Script properties]';
}
var LOG_FOLDER_ID = PropertiesService.getScriptProperties().getProperty('log_folder_id');
if (!LOG_FOLDER_ID) {
    throw 'You should set "log_folder_id" property from [File] > [Project properties] > [Script properties]';
}
var LOG_LEVELS = ["trace", "debug", "info", "warn", "error", "fatal"];
var LOG_LEVEL = LOG_LEVELS.indexOf("info");
var LOG_LEVEL_PROP = PropertiesService.getScriptProperties().getProperty('log_level');
if (LOG_LEVELS.indexOf(LOG_LEVEL_PROP) >= 0) {
    LOG_LEVEL = LOG_LEVELS.indexOf(LOG_LEVEL_PROP);
}
var SpreadsheetLogger = (function () {
    function SpreadsheetLogger(id) {
        this.id = id;
    }
    SpreadsheetLogger.prototype.log_sheet_ = function () {
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
    };
    SpreadsheetLogger.prototype.log_ = function (level, message) {
        if (LOG_LEVELS.indexOf(level) < LOG_LEVEL) {
            return;
        }
        var sh = this.log_sheet_();
        var now = new Date();
        var last_row = sh.getLastRow();
        sh.insertRowAfter(last_row).getRange(last_row + 1, 1, 1, 3).setValues([[now, level, "'" + message]]);
        return sh;
    };
    SpreadsheetLogger.prototype.trace = function (message) {
        this.log_('trace', message);
    };
    SpreadsheetLogger.prototype.debug = function (message) {
        this.log_('debug', message);
    };
    SpreadsheetLogger.prototype.info = function (message) {
        this.log_('info', message);
    };
    SpreadsheetLogger.prototype.warn = function (message) {
        this.log_('warn', message);
    };
    SpreadsheetLogger.prototype.error = function (message) {
        this.log_('error', message);
    };
    SpreadsheetLogger.prototype.fatal = function (message) {
        this.log_('fatal', message);
    };
    return SpreadsheetLogger;
}());
var myLogger = new SpreadsheetLogger(APP_WORKSHEET_ID);
var ChannelLogStatus = (function () {
    function ChannelLogStatus(key, timestamp, status) {
        this.key = key;
        this.timestamp = timestamp;
        this.status = status;
    }
    ChannelLogStatus.prototype.toObjectArray = function () {
        return [this.key, this.timestamp, this.status];
    };
    return ChannelLogStatus;
}());
var SpreadsheetKeyValueStore = (function () {
    function SpreadsheetKeyValueStore(id) {
        this.id = id;
        this.keyStatusMap = {};
        this.init();
    }
    SpreadsheetKeyValueStore.prototype.init = function () {
        var _this = this;
        var sh = this.getSheet();
        var values = sh.getSheetValues(1, 1, sh.getMaxRows() - 1, sh.getMaxColumns());
        values.forEach(function (v, i) {
            var status = new ChannelLogStatus(v[0], v[1], v[2]);
            _this.keyStatusMap[status.key] = { index: i + 1, values: status };
        });
        return sh;
    };
    SpreadsheetKeyValueStore.prototype.getSheet = function () {
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
    };
    SpreadsheetKeyValueStore.prototype.newRow = function (key) {
        var sh = this.getSheet();
        var last_row = sh.getLastRow();
        sh.insertRowAfter(last_row);
        this.keyStatusMap[key] = { index: last_row + 1, values: null };
        return last_row + 1;
    };
    SpreadsheetKeyValueStore.prototype.setStatus = function (key, status) {
        if (!this.keyStatusMap[key]) {
            this.newRow(key);
        }
        var keyInfo = this.keyStatusMap[key];
        var sh = this.getSheet();
        var now = new Date();
        keyInfo.values = new ChannelLogStatus(key, now, status);
        sh.getRange(keyInfo.index, 1, 1, 3).setValues([keyInfo.values.toObjectArray()]).clearFormat();
    };
    SpreadsheetKeyValueStore.prototype.getStatus = function (key) {
        if (this.keyStatusMap[key]) {
            return this.keyStatusMap[key].values;
        }
        return new ChannelLogStatus(key, new Date(1), "");
    };
    return SpreadsheetKeyValueStore;
}());
var keyValueStore = new SpreadsheetKeyValueStore(APP_WORKSHEET_ID);
function StoreLogsDelta() {
    var logger = new SlackChannelHistoryLogger();
    myLogger.info("Start logger.run");
    logger.run();
    myLogger.info("End   logger.run");
}
;
var SlackChannelHistoryLogger = (function () {
    function SlackChannelHistoryLogger() {
        this.memberNames = {};
        this.cachedSpreadSheet = {};
        this.cachedSheet = {};
    }
    SlackChannelHistoryLogger.prototype.requestSlackAPI = function (path, params) {
        if (params === void 0) { params = {}; }
        var url = "https://slack.com/api/" + path + "?";
        var qparams = [("token=" + encodeURIComponent(API_TOKEN))];
        for (var k in params) {
            qparams.push(encodeURIComponent(k) + "=" + encodeURIComponent(params[k]));
        }
        url += qparams.join('&');
        myLogger.debug("==> GET " + url);
        var resp = UrlFetchApp.fetch(url);
        var data = JSON.parse(resp.getContentText());
        if (data.error) {
            throw "GET " + path + ": " + data.error;
        }
        myLogger.debug("<== GOT");
        return data;
    };
    SlackChannelHistoryLogger.prototype.run = function () {
        var _this = this;
        //時刻格納用の変数
        var starttime = +new Date();
        var usersResp = this.requestSlackAPI('users.list');
        usersResp.members.forEach(function (member) {
            _this.memberNames[member.id] = member.name;
        });
        var teamInfoResp = this.requestSlackAPI('team.info');
        this.teamName = teamInfoResp.team.name;
        var channelsResp = this.requestSlackAPI('channels.list');
        var channelFetchTime = function (ch) {
            var sheetName = _this.sheetName(ch);
            var status = keyValueStore.getStatus(sheetName);
            var time = status.timestamp.getTime();
            if (status.status != FETCH_STATUS_END) {
                time = 0;
            }
            return time;
        };
        channelsResp.channels.sort(function (ch1, ch2) {
            var time1 = channelFetchTime(ch1);
            var time2 = channelFetchTime(ch2);
            return time1 - time2;
        });
        for (var _i = 0, _a = channelsResp.channels; _i < _a.length; _i++) {
            var ch = _a[_i];
            this.importChannelHistoryDelta(ch);
            var endtime = +new Date();
            if (endtime - starttime > TRIGGER_LIMIT) {
                myLogger.warn("TERMINATE by limit time " + (endtime - starttime) + " > " + TRIGGER_LIMIT);
                break;
            }
        }
    };
    SlackChannelHistoryLogger.prototype.getLogsFolder = function () {
        if (!this.cachedFolder) {
            var folder = DriveApp.getFolderById(LOG_FOLDER_ID);
            if (!folder) {
                throw 'You should set "log_folder_id" property from [File] > [Project properties] > [Script properties]';
            }
            this.cachedFolder = folder;
        }
        return this.cachedFolder;
    };
    SlackChannelHistoryLogger.prototype.getSpreadSheet = function (ch, d, readonly) {
        if (readonly === void 0) { readonly = false; }
        var dateString;
        if (d instanceof Date) {
            dateString = this.formatDate(d);
        }
        else {
            dateString = '' + d;
        }
        var spreadsheet;
        var spreadsheetName = dateString;
        if (this.cachedSpreadSheet[spreadsheetName]) {
            spreadsheet = this.cachedSpreadSheet[spreadsheetName];
        }
        else {
            var folder = this.getLogsFolder();
            var it = folder.getFilesByName(spreadsheetName);
            if (it.hasNext()) {
                var file = it.next();
                spreadsheet = SpreadsheetApp.openById(file.getId());
            }
            else {
                if (readonly)
                    return null;
                spreadsheet = SpreadsheetApp.create(spreadsheetName);
                folder.addFile(DriveApp.getFileById(spreadsheet.getId()));
            }
        }
        return spreadsheet;
    };
    SlackChannelHistoryLogger.prototype.sheetName = function (ch) {
        var sheetName = ch.name + " (" + ch.id + ")";
        return sheetName;
    };
    SlackChannelHistoryLogger.prototype.getSheet = function (ch, d, readonly) {
        if (readonly === void 0) { readonly = false; }
        var spreadsheet = this.getSpreadSheet(ch, d, readonly);
        if (!spreadsheet) {
            return null;
        }
        var dateString;
        if (d instanceof Date) {
            dateString = this.formatDate(d);
        }
        else {
            dateString = '' + d;
        }
        var sheetByID;
        if (this.cachedSheet[dateString]) {
            sheetByID = this.cachedSheet[dateString];
        }
        else {
            sheetByID = {};
            var sheets = spreadsheet.getSheets();
            sheets.forEach(function (s) {
                var name = s.getName();
                var m = /^(.+) \((.+)\)$/.exec(name); // eg. "general (C123456)"
                if (!m)
                    return;
                sheetByID[m[2]] = s;
            });
            this.cachedSheet[dateString] = sheetByID;
        }
        var sheet = sheetByID[ch.id];
        if (!sheet) {
            if (readonly)
                return null;
            sheet = spreadsheet.insertSheet();
            sheet.setColumnWidth(COL_LOG_TEXT, 800);
        }
        var sheetName = this.sheetName(ch);
        if (sheet.getName() !== sheetName) {
            sheet.setName(sheetName);
        }
        return sheet;
    };
    SlackChannelHistoryLogger.prototype.importChannelHistoryDelta = function (ch) {
        var _this = this;
        myLogger.info("importChannelHistoryDelta " + ch.name + " (" + ch.id + ")");
        var sheetName = this.sheetName(ch);
        var prevStatus = keyValueStore.getStatus(sheetName);
        if (prevStatus.status == FETCH_STATUS_ARCHIVED) {
            return;
        }
        keyValueStore.setStatus(sheetName, FETCH_STATUS_START);
        var now = new Date();
        var oldest = '1'; // oldest=0 does not work
        var existingSheet = this.getSheet(ch, now, true);
        if (!existingSheet) {
            // try previous month
            now.setMonth(now.getMonth() - 1);
            existingSheet = this.getSheet(ch, now, true);
        }
        if (existingSheet) {
            var lastRow = existingSheet.getLastRow();
            try {
                var data = JSON.parse(existingSheet.getRange(lastRow, COL_LOG_RAW_JSON).getValue());
                oldest = data.ts;
            }
            catch (e) {
                myLogger.warn("while trying to parse the latest history item from existing sheet: " + e);
            }
        }
        var messages = this.loadMessagesBulk(ch, { oldest: oldest });
        var dateStringToMessages = {};
        messages.forEach(function (msg) {
            var date = new Date(+msg.ts * 1000);
            var dateString = _this.formatDate(date);
            if (!dateStringToMessages[dateString]) {
                dateStringToMessages[dateString] = [];
            }
            dateStringToMessages[dateString].push(msg);
        });
        for (var dateString in dateStringToMessages) {
            var sheet = this.getSheet(ch, dateString);
            var timezone = sheet.getParent().getSpreadsheetTimeZone();
            var lastTS = 0;
            var lastRow = sheet.getLastRow();
            if (lastRow > 0) {
                try {
                    var data = JSON.parse(sheet.getRange(lastRow, COL_LOG_RAW_JSON).getValue());
                    lastTS = +data.ts || 0;
                }
                catch (_) {
                }
            }
            var rows = dateStringToMessages[dateString].filter(function (msg) {
                return +msg.ts > lastTS;
            }).map(function (msg) {
                var date = new Date(+msg.ts * 1000);
                return [
                    Utilities.formatDate(date, timezone, 'yyyy-MM-dd HH:mm:ss'),
                    _this.memberNames[msg.user] || msg.username,
                    _this.unescapeMessageText(msg.text),
                    JSON.stringify(msg)
                ];
            });
            if (rows.length > 0) {
                var range = sheet.insertRowsAfter(lastRow || 1, rows.length)
                    .getRange(lastRow + 1, 1, rows.length, COL_MAX);
                range.setValues(rows);
            }
        }
        if (ch.is_archived) {
            keyValueStore.setStatus(sheetName, FETCH_STATUS_ARCHIVED);
        }
        else {
            keyValueStore.setStatus(sheetName, FETCH_STATUS_END);
        }
    };
    SlackChannelHistoryLogger.prototype.formatDate = function (dt) {
        return Utilities.formatDate(dt, Session.getScriptTimeZone(), 'yyyy-MM');
    };
    SlackChannelHistoryLogger.prototype.loadMessagesBulk = function (ch, options) {
        var _this = this;
        if (options === void 0) { options = {}; }
        var messages = [];
        // channels.history will return the history from the latest to the oldest.
        // If the result's "has_more" is true, the channel has more older history.
        // In this case, use the result's "latest" value to the channel.history API parameters
        // to obtain the older page, and so on.
        options['count'] = HISTORY_COUNT_PER_PAGE;
        options['channel'] = ch.id;
        var loadSince = function (oldest) {
            if (oldest) {
                options['oldest'] = oldest;
            }
            // order: recent-to-older
            var resp = _this.requestSlackAPI('channels.history', options);
            messages = resp.messages.concat(messages);
            return resp;
        };
        var resp = loadSince();
        var page = 1;
        while (resp.has_more && page <= MAX_HISTORY_PAGINATION) {
            myLogger.info("channels.history.pagination " + ch.name + " (" + ch.id + ") " + page);
            resp = loadSince(resp.messages[0].ts);
            page++;
        }
        // oldest-to-recent
        return messages.reverse();
    };
    SlackChannelHistoryLogger.prototype.unescapeMessageText = function (text) {
        var _this = this;
        return (text || '')
            .replace(/&lt;/g, '<')
            .replace(/&gt;/g, '>')
            .replace(/&quot;/g, '"')
            .replace(/&amp;/g, '&')
            .replace(/<@(.+?)>/g, function ($0, userID) {
            var name = _this.memberNames[userID];
            return name ? "@" + name : $0;
        });
    };
    return SlackChannelHistoryLogger;
}());
