/**
 * ファイルを開いたときのイベントハンドラ
 */
function onOpen() {
    var ui = SpreadsheetApp.getUi();           // Uiクラスを取得する
    var menu = ui.createMenu('書き出し');  // Uiクラスからメニューを作成する
    menu.addItem('計画書き出し', 'plan');   // メニューにアイテムを追加する
    menu.addItem('実働書き出し', 'act');   // メニューにアイテムを追加する
    menu.addToUi();                            // メニューをUiクラスに追加する
}

function plan() {
    let proc = new TableProc();
    proc.writePlan();
}

function act() {
    let proc = new TableProc();
    proc.writeAct();
}

class TableProc {

    // 計画出力
    writePlan() {
        let sheet = SpreadsheetApp.getActive();

        let nmStack: string[] = []; // 慣例
        let bgStack: string[] = []; // 大項目
        let mdStack: string[] = []; // 中項目
        let smStack: string[] = []; // 小項目
        let tmStack: Date[] = []; // 慣例
        let from = this.toN('E');
        let to = this.toN('DH');

        let recIndex = 0;
        for (let i = from; i <= to; i++) {
            let a = this.toA(i);
            let nm = sheet.getRange(a + 4).getValue().toString();
            let bg = sheet.getRange(a + 6).getValue().toString();
            let md = sheet.getRange(a + 7).getValue().toString();
            let sm = sheet.getRange(a + 8).getValue().toString();
            if (nm != "" ||
                bg != "" ||
                md != "" ||
                sm != "") {
                tmStack[recIndex] = new Date(Date.UTC(0, 0, 0, 15, (i - from) * 5, 0));
                nmStack[recIndex] = nm;
                bgStack[recIndex] = bg;
                mdStack[recIndex] = md;
                smStack[recIndex] = sm;
                recIndex++;
            }
        }
        // 最後の日付追加
        const START = 12;
        for (let i = 0; i < nmStack.length; i++) {
            if (nmStack[i] != "" ||
                bgStack[i] != "" ||
                mdStack[i] != "" ||
                smStack[i] != "") {
                let tm = tmStack[i];
                let tmN = tmStack[i + 1];
                if (tmN === undefined) {
                    tmN = new Date(Date.UTC(0, 0, 0, 24, 0, 0, 0));
                }
                let dv = new Date(Date.UTC(0, 0, 0, 5, 0, 0, tmN.getTime() - tm.getTime()));
                let nm = nmStack[i];
                let bg = bgStack[i];
                let md = mdStack[i];
                let sm = smStack[i];
                if (nm != "業務" && bg == "" && md == "" && sm == "") {
                    sheet.getRange("Z" + (i + START)).setValue(nm);
                    sheet.getRange("AA" + (i + START)).setValue("");
                    sheet.getRange("AB" + (i + START)).setValue("");
                } else if (nm == "業務" && bg == "" && md == "" && sm == "") {
                    // 一つ前を取得
                    let prevBg = null;
                    let prevMd = null;
                    let prevSm = null;
                    for (let j = i - 1; j >= 0; j--) {
                        if (prevBg == null && bgStack[j] != "") {
                            prevBg = bgStack[j];
                        }
                        if (prevMd == null && mdStack[j] != "") {
                            prevMd = mdStack[j];
                        }
                        if (prevSm == null && smStack[j] != "") {
                            prevSm = smStack[j];
                        }
                        if (prevBg != null) {
                            break;
                        }
                    }
                    if (bg == "" && (md != "" || sm != "")) {
                        bg = "↓";
                    }
                    if (md == "" && sm != "") {
                        md = "↓";
                    }

                    sheet.getRange("Z" + (i + START)).setValue(prevBg);
                    sheet.getRange("AA" + (i + START)).setValue(prevMd);
                    sheet.getRange("AB" + (i + START)).setValue(prevSm);
                } else {
                    if (bg == "" && (md != "" || sm != "")) {
                        bg = "↓";
                    }
                    if (md == "" && sm != "") {
                        md = "↓";
                    }
                    sheet.getRange("Z" + (i + START)).setValue(bg);
                    sheet.getRange("AA" + (i + START)).setValue(md);
                    sheet.getRange("AB" + (i + START)).setValue(sm);
                }

                sheet.getRange("AC" + (i + START)).setValue("");
                sheet.getRange("AD" + (i + START)).setValue("");
                sheet.getRange("AE" + (i + START)).setValue(this.format(dv, "hh:mm"));
                sheet.getRange("AF" + (i + START)).setValue(this.format(tm, "hh:mm"));
                sheet.getRange("AG" + (i + START)).setValue(this.format(tmN, "hh:mm"));
                sheet.getRange("AH" + (i + START)).setValue(this.format(dv, "hh:mm"));
            }
        }
    }

    // 計画出力
    writeAct() {
        let sheet = SpreadsheetApp.getActive();

        let smStack: string[] = []; // 小項目
        let tmStack: Date[] = []; // 慣例
        let from = this.toN('E');
        let to = this.toN('DH');

        let recIndex = 0;
        for (let i = from; i <= to; i++) {
            let a = this.toA(i);
            let sm = sheet.getRange(a + 10).getValue().toString();
            if (sm != "") {
                tmStack[recIndex] = new Date(Date.UTC(0, 0, 0, 15, (i - from) * 5, 0));
                smStack[recIndex] = sm;
                recIndex++;
            }
        }
        // 最後の日付追加
        const START = 12;
        for (let i = 0; i < smStack.length; i++) {
            if (smStack[i] != "") {
                let tm = tmStack[i];
                let tmN = tmStack[i + 1];
                if (tmN === undefined) {
                    tmN = new Date(Date.UTC(0, 0, 0, 24, 0, 0, 0));
                }
                let dv = new Date(Date.UTC(0, 0, 0, 5, 0, 0, tmN.getTime() - tm.getTime()));
                let sm = smStack[i];

                sheet.getRange("AN" + (i + START)).setValue(sm);
                // 隙間
                sheet.getRange("AX" + (i + START)).setValue(this.format(tm, "hh:mm"));
                sheet.getRange("AY" + (i + START)).setValue(this.format(tmN, "hh:mm"));
                sheet.getRange("AZ" + (i + START)).setValue(this.format(dv, "hh:mm"));
            }
        }
    }
    toA(column: number): string {
        let temp, letter = '';
        while (column > 0) {
            temp = (column - 1) % 26;
            letter = String.fromCharCode(temp + 65) + letter;
            column = (column - temp - 1) / 26;
        }
        return letter;
    }

    toN(letter: string): number {
        let column = 0, length = letter.length;
        for (let i = 0; i < length; i++) {
            column += (letter.charCodeAt(i) - 64) * Math.pow(26, length - i - 1);
        }
        return column;
    }

    // フォーマット
    format(date, format) {
        if (!format) format = 'YYYY-MM-DD hh:mm:ss.SSS';
        format = format.replace(/YYYY/g, date.getFullYear());
        format = format.replace(/MM/g, ('0' + (date.getMonth() + 1)).slice(-2));
        format = format.replace(/DD/g, ('0' + date.getDate()).slice(-2));
        format = format.replace(/hh/g, ('0' + date.getHours()).slice(-2));
        format = format.replace(/mm/g, ('0' + date.getMinutes()).slice(-2));
        format = format.replace(/ss/g, ('0' + date.getSeconds()).slice(-2));
        if (format.match(/S/g)) {
            var milliSeconds = ('00' + date.getMilliseconds()).slice(-3);
            var length = format.match(/S/g).length;
            for (var i = 0; i < length; i++) format = format.replace(/S/, milliSeconds.substring(i, i + 1));
        }
        return format;
    }
}