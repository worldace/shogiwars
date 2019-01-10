/* ※このファイルはSJISで保存してください */



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
    var 解析結果   = [];

    for(var i = 0; i < a.length; i++){
       if(a[i].textContent === '\u898B\u308B'){ //見る
            var 棋譜ID   = String(a[i].onclick).match(/'(.+?)'/)[1];
            var ファイル = 将棋ウォーズ.棋譜IDをファイル名に変換(棋譜ID)
            if(将棋ウォーズ.ダウンロード済み棋譜一覧.indexOf(ファイル) !== -1){
                break;
            }
            解析結果.push(棋譜ID);
        }
        else if(a[i].textContent === '\u6B21'){ //次
            解析結果.次のページ = a[i].href;
            break;
        }
    }

    for(var i = 0; i < 解析結果.length; i++){
        var 棋譜ページ解析結果 = 将棋ウォーズ.棋譜ページ解析(解析結果[i]);
        if(棋譜ページ解析結果.棋譜){
            将棋ウォーズ.棋譜ファイル保存(棋譜ページ解析結果);
            将棋ウォーズ.取得件数++;
        }
    }
    if(解析結果.次のページ){
        将棋ウォーズ.棋譜一覧ページ解析(解析結果.次のページ);
    }
};



将棋ウォーズ.棋譜ページ解析 = function(棋譜ID){
    var html   = HTTP_GET("https://kif-pona.heroz.jp/games/" + 棋譜ID);
    var ソース = html.substr(html.indexOf('gamedata'));

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
    kif += "開始日時：" + 将棋ウォーズ.棋譜IDを時間に変換(解析結果.棋譜ID) + "\r\n";
    kif += "先手：" + 解析結果.先手名前 + " "+ 解析結果.先手段級 + "\r\n";
    kif += "後手：" + 解析結果.後手名前 + " "+ 解析結果.後手段級 + "\r\n";
    kif += "棋戦：将棋ウォーズ " + 将棋ウォーズ.現在のモード + "\r\n";
    kif += "場所：https://kif-pona.heroz.jp/games/" + 解析結果.棋譜ID + "\r\n";
    kif +=  将棋ウォーズ.KIF変換(解析結果.棋譜);

    var パス = 将棋ウォーズ.現在のモード + '/' + 将棋ウォーズ.棋譜IDをファイル名に変換(解析結果.棋譜ID) + '.kif';
    ファイル保存(パス, kif, "Shift_JIS");
};



将棋ウォーズ.KIF変換 = function(csa){
    var result = "";
    var 駒変換 = {FU:'歩', KY:'香' ,KE:'桂' ,GI:'銀' ,KI:'金', KA:'角', HI:'飛', OU:'玉', TO:'と', NY:'成香' ,NK:'成桂', NG:'成銀', UM:'馬', RY:'龍'};
    var 逆変換 = {TO:'歩', NY:'香' ,NK:'桂', NG:'銀', UM:'角', RY:'飛'};
    var 成り駒 = ['TO', 'NY', 'NK', 'NG', 'UM', 'RY'];
    var 全数字 = {1:'１', 2:'２', 3:'３', 4:'４', 5:'５', 6:'６', 7:'７', 8:'８', 9:'９'};
    var 漢数字 = {1:'一', 2:'二', 3:'三', 4:'四', 5:'五', 6:'六', 7:'七', 8:'八', 9:'九'};
    var 終局   = {CHECKMATE:'詰み', TORYO:'投了', DISCONNECT:'反則負け', TIMEOUT:'切れ負け', OUTE_SENNICHI:'反則負け', SENNICHI:'千日手'};
    var 時間表 = {'10分':600, '3分':180, '10秒':3600};

    var 先手残り時間 = 時間表[将棋ウォーズ.現在のモード];
    var 後手残り時間 = 時間表[将棋ウォーズ.現在のモード];
    var 先手累積時間 = 0;
    var 後手累積時間 = 0;

    csa = csa.split("\t");
    for(var i = 0; i < csa.length; i++){
        var 手数   = i + 1;
        var 指し手 = csa[i].split(",")[0];
        var 時間   = csa[i].split(",")[1];

        if(指し手.match(/^[\+\-]/)){
            var kif  = "";
            var 手番 = 指し手.substr(0, 1);
            var 前X  = 指し手.substr(1, 1);
            var 前Y  = 指し手.substr(2, 1);
            var 後X  = 指し手.substr(3, 1);
            var 後Y  = 指し手.substr(4, 1);
            var 駒   = 指し手.substr(5, 2);
            
            kif += 全数字[後X]; //"同　歩"みたいな表現に対応していない
            kif += 漢数字[後Y];

            if(成り駒.indexOf(駒) > -1 && 将棋ウォーズ.KIF変換.成り判定(csa, i, 手番, 駒, 前X, 前Y)){
                kif += 逆変換[駒] + "成";
            }
            else{
                kif += 駒変換[駒];
            }
            kif += (前X == 0) ? '打' : ('(' + 前X + 前Y + ')');

            //時間
            var 時間 = Number(時間.substr(1));
            if(手番 === "+"){
                var 消費時間  = 先手残り時間 - 時間;
                先手累積時間 += 消費時間;
                先手残り時間  = 時間;
                kif          += "  " + 将棋ウォーズ.KIF変換.時間(消費時間, 先手累積時間);
            }
            else{
                var 消費時間  = 後手残り時間 - 時間;
                後手累積時間 += 消費時間;
                後手残り時間  = 時間;
                kif          += "  " + 将棋ウォーズ.KIF変換.時間(消費時間, 後手累積時間);
            }
            result += 手数 + " " + kif + "\r\n";
        }
        else{
            var 終局文字列 = 指し手.match(/(_WIN|DRAW)_(\w+)/);
            if(終局文字列[2] in 終局){
                result += 手数 + " " + 終局[終局文字列[2]] + "\r\n";
            }
            break;
        }
    }

    return result;
};



将棋ウォーズ.KIF変換.成り判定 = function (csa, 手数, 手番, 成り駒, 前X, 前Y){
    //前回の位置が成り駒であるとfalse (それ以外はtrue)
    //検索対象は現在の手数まで。自分の手番のみ
    var 判定 = new RegExp('^\\' + 手番 + '\\d\\d' + 前X + 前Y);

    for(var i = 手数; i >= 0; i -= 2){
        if(csa[i].match(判定)){
            return (csa[i].match(成り駒)) ? false : true;
        }
    }
    return true;
};



将棋ウォーズ.KIF変換.時間 = function (消費時間, 累積時間){
    var 消費分 = Math.floor(消費時間 / 60);
    var 消費秒 = 消費時間 % 60;
    var 累積時 = Math.floor(累積時間 / 3600);
    var 累積分 = Math.floor(累積時間 % 3600 / 60);
    var 累積秒 = 累積時間 % 60;

    消費秒 = ('0' + 消費秒).slice(-2);
    累積時 = ('0' + 累積時).slice(-2);
    累積分 = ('0' + 累積分).slice(-2);
    累積秒 = ('0' + 累積秒).slice(-2);

    return "( " + 消費分 + ":" + 消費秒 + "/" + 累積時 + ":" + 累積分 + ":" + 累積秒 + ")";
};


将棋ウォーズ.棋譜IDをファイル名に変換 = function (棋譜ID){
    var id = 棋譜ID.split('-');
    return id[2] + "-" + id[0] + "-" + id[1];
}



将棋ウォーズ.棋譜IDを時間に変換 = function (棋譜ID){
    var id = 棋譜ID.split('-');

    var 年 = id[2].substr(0, 4);
    var 月 = id[2].substr(4, 2);
    var 日 = id[2].substr(6, 2);
    var 時 = id[2].substr(9, 2);
    var 分 = id[2].substr(11, 2);
    var 秒 = id[2].substr(13, 2);

    return 年 + "/" + 月 + "/" + 日 + " " + 時 + ":" + 分 + ":" + 秒;
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



function HTTP_GET(url){
    var http = new ActiveXObject('Msxml2.ServerXMLHTTP');
    http.open('GET', url, false);
    http.send();
    return http.responseText;
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


var IE = new ActiveXObject("InternetExplorer.Application");
IE.visible = false;
将棋ウォーズ();


// ウォーズ仕様の参考ページ https://github.com/tosh1ki/shogiwars
// CSA形式の参考ページ http://www2.computer-shogi.org/protocol/record_v22.html
// KIF形式の参考ページ http://kakinoki.o.oo7.jp/kif_format.html