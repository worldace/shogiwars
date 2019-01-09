/* ※このファイルはSJISで保存してください */

var IE = new ActiveXObject("InternetExplorer.Application");
IE.visible = false;


function 将棋ウォーズ(){
    多重起動防止();
    WScript.Echo("棋譜のダウンロードを開始します");

    将棋ウォーズ.取得件数 = 0;
    将棋ウォーズ.ユーザID = ファイル取得("id.txt").replace("/\s/g", "");
    if(!将棋ウォーズ.ユーザID){
        将棋ウォーズ.終了("id.txtに将棋ウォーズのIDを記述してください")
    }

    var gtype = [{mode: '10分', name: ''}, {mode: '3分', name: 'sb'}, {mode: '10秒', name: 's1'}];

    for(var i = 0; i < gtype.length; i++){
        将棋ウォーズ.現在のモード = gtype[i].mode;
        将棋ウォーズ.ダウンロード済み棋譜一覧 = ファイル一覧(gtype[i].mode, "base");
        将棋ウォーズ.棋譜一覧ページ解析("https://shogiwars.heroz.jp/games/history?gtype=" + gtype[i].name + "&user_id=" + 将棋ウォーズ.ユーザID);
    }

    将棋ウォーズ.終了(将棋ウォーズ.取得件数 + "件のファイルを取得しました")
}



将棋ウォーズ.棋譜一覧ページ解析 = function(url){
    var document   = IE移動(url);
    var a          = document.querySelectorAll("a");

    for(var i = 0; i < a.length; i++){
        if(a[i].textContent === '\u898B\u308B'){ //見る
            var 棋譜ID = String(a[i].onclick).match(/'(.+?)'/)[1];
            if(将棋ウォーズ.ダウンロード済み棋譜一覧.indexOf(棋譜ID) !== -1){
                return;
            }

            var 棋譜ページ解析結果 = 将棋ウォーズ.棋譜ページ解析(棋譜ID);
            if(棋譜ページ解析結果.棋譜){
                将棋ウォーズ.棋譜ファイル保存(棋譜ページ解析結果);
                将棋ウォーズ.取得件数++;
                return;
            }
        }
        else if(a[i].textContent === '\u6B21'){ //次
            将棋ウォーズ.棋譜一覧ページ解析(a[i].href);
            break;
        }
    }
};


将棋ウォーズ.棋譜ページ解析 = function(棋譜ID){
    var document = IE移動("https://kif-pona.heroz.jp/games/" + 棋譜ID);
    var script   = document.querySelectorAll("script");
    var ソース   = script[script.length-3].textContent; //決め打ち

    return {
        棋譜ID  : 棋譜ID,
        先手名前: 棋譜ID.split(/-/)[0],
        後手名前: 棋譜ID.split(/-/)[1],
        先手段級: ソース.match(/dan0: "(.+?)"/)[1],
        後手段級: ソース.match(/dan1: "(.+?)"/)[1],
        棋譜    : ソース.match(/receiveMove\("(.+?)"/)[1]
    };
};



将棋ウォーズ.棋譜ファイル保存 = function(解析結果){
    var kif = "";
    kif += "先手：" + 解析結果.先手名前 + " "+ 解析結果.先手段級 + "\r\n";
    kif += "後手：" + 解析結果.後手名前 + " "+ 解析結果.後手段級 + "\r\n";
    kif += "場所：https://kif-pona.heroz.jp/games/" + 解析結果.棋譜ID + "\r\n";
    kif +=  将棋ウォーズ.KIF変換(解析結果.棋譜);

    var パス   = 将棋ウォーズ.現在のモード + '/' + 解析結果.棋譜ID + '.kif';
    ファイル保存(パス, kif, "Shift_JIS");
};



将棋ウォーズ.KIF変換 = function(csa){
    var result = "";
    var 駒変換 = {FU:'歩', KY:'香' ,KE:'桂' ,GI:'銀' ,KI:'金', KA:'角', HI:'飛', OU:'玉', TO:'と', NY:'成香' ,NK:'成桂', NG:'成銀', UM:'馬', RY:'龍'};
    var 逆変換 = {TO:'歩', NY:'香' ,NK:'桂', NG:'銀', UM:'角', RY:'飛'};
    var 成り駒 = ['TO', 'NY', 'NK', 'NG', 'UM', 'RY'];
    var 全数字 = {1:'１', 2:'２', 3:'３', 4:'４', 5:'５', 6:'６', 7:'７', 8:'８', 9:'９'};
    var 漢数字 = {1:'一', 2:'二', 3:'三', 4:'四', 5:'五', 6:'六', 7:'七', 8:'八', 9:'九'};

    csa = csa.split("\t");
    for(var i = 0; i < csa.length; i++){
        var kif    = (i + 1) + " ";
        var 指し手 = csa[i].split(",")[0];
        var 時間   = csa[i].split(",")[1];

        if(指し手.match(/^[\+\-]/)){
            var 前X = 指し手.substr(1, 1);
            var 前Y = 指し手.substr(2, 1);
            var 後X = 指し手.substr(3, 1);
            var 後Y = 指し手.substr(4, 1);
            var 駒  = 指し手.substr(5, 2);
            
            kif += 全数字[後X]; //"同　歩"みたいな表現に対応していない
            kif += 漢数字[後Y];

            if(成り駒.indexOf(駒) > -1 && 将棋ウォーズ.KIF変換.成り判定(csa, i, 駒, 前X, 前Y)){
                kif += 逆変換[駒] + "成";
            }
            else{
                kif += 駒変換[駒];
            }
            kif += (前X == 0) ? '打' : ('(' + 前X + 前Y + ')');
        }
        else{
            
        }
        result += kif + "\r\n";
    }

    return result;
};


将棋ウォーズ.KIF変換.成り判定 = function (csa, 手数, 成り駒, 前X, 前Y){
    //前回の位置が成り駒であるとfalse (それ以外はtrue)
    //検索対象は現在の手数まで。自分の手番のみ
    var 判定 = 前X + 前Y + 成り駒;
    var 手番 = 手数 % 2;

    for(var i = 0; i < 手数; i++){
        if(i % 2 !== 手番){
            continue;
        }
        if(csa[i].match(判定)){
            return false;
        }
    }
    return true;
};



将棋ウォーズ.終了 = function(str){
    IE.Quit();
    WScript.Echo(str);
    WScript.Quit();
};



function IE移動(url){ /* エラー時の対策が分からない */
    IE.navigate(url);
    for(var i = 0; i < 300; i++){
        if(IE.Busy || IE.readystate !== 4){
            WScript.Sleep(100);
            continue;
        }
        break;
    }
    return IE.document;
}



function ファイル取得(file, encode){
    var stream = new ActiveXObject('ADODB.Stream');
    stream.charset = encode || 'utf-8';
    stream.Open();
    stream.loadFromFile(file);
    var contents = stream.ReadText();
    stream.close();

    return contents;
}



function ファイル保存(file, contents, encode){
    var stream = new ActiveXObject('ADODB.Stream');
    stream.charset = encode || 'utf-8';
    stream.Open();
    stream.WriteText(contents);
    stream.SaveToFile(file, 1);
    stream.Close();
}



function ファイル一覧(dir, type){
    var result = [];
    var fs     = new ActiveXObject("Scripting.FileSystemObject");
    var files  = fs.GetFolder(dir).Files;

    for (var e = new Enumerator(files); !e.atEnd(); e.moveNext()){
        if(type === 'base'){
            result.push(fs.GetBaseName(e.item()));
        }
        else if(type === 'absolute'){
            result.push(e.item());
        }
        else{
            result.push(fs.GetFileName(e.item()));
        }
    }
    return result;
}




function 多重起動防止(){ // https://gist.github.com/ka-ka-xyz/2718628
    var env = WScript.CreateObject("WScript.Shell").Environment("Process");
    var windir = env("SystemRoot");
    var wmiObj = GetObject("WinMgmts:Root\\Cimv2");
    var processes = wmiObj.ExecQuery("Select * From Win32_Process");
    var processenum = new Enumerator(processes);
    var flag = 0;
    var pid;

    for(; !processenum.atEnd(); processenum.moveNext()){
        var item = processenum.item();
        if(item.CommandLine != null && item.CommandLine.toLowerCase().indexOf("\"" +  windir.toLowerCase() + "\\system32\\wscript.exe") == 0 && item.CommandLine.indexOf(WScript.ScriptFullName) > 0 ){
            flag++;
            pid = item.ProcessId;
        }
    }
  
    if(flag > 1){
        WScript.Echo(WScript.ScriptName + "は既に起動されています。");
        WScript.Quit(0);
    }
    else if(typeof pid === "undefined"){
        WScript.Echo("多重起動判定に失敗しました。");
        WScript.Quit(0);
    }
    return pid;
}



Array.prototype.indexOf = function(obj, start){
    for(var i = (start || 0), j = this.length; i < j; i++){
        if(this[i] === obj){
            return i;
        }
    }
    return -1;
};



将棋ウォーズ();


// ウォーズ仕様の参考ページ
// https://github.com/teriyaki398/MyKifu
// https://github.com/tosh1ki/shogiwars
