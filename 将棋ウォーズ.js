/* �����̃t�@�C����SJIS�ŕۑ����Ă������� */



function �����E�H�[�Y(){
    ���d�N���h�~();
    WScript.Echo("�����̃_�E�����[�h���J�n���܂�");

    �����E�H�[�Y.�擾���� = 0;
    �����E�H�[�Y.���[�UID = �t�@�C���擾("id.txt").replace("/\s/g", "");
    if(!�����E�H�[�Y.���[�UID){
        �����E�H�[�Y.�I��("id.txt�ɏ����E�H�[�Y��ID���L�q���Ă�������")
    }

    var gtype = [{mode: '10��', name: ''}, {mode: '3��', name: 'sb'}, {mode: '10�b', name: 's1'}];

    for(var i = 0; i < gtype.length; i++){
        �����E�H�[�Y.���݂̃��[�h = gtype[i].mode;
        �����E�H�[�Y.�_�E�����[�h�ς݊����ꗗ = �t�@�C���ꗗ(gtype[i].mode, "base");

        �����E�H�[�Y.�����ꗗ�y�[�W���("https://shogiwars.heroz.jp/games/history?gtype=" + gtype[i].name + "&user_id=" + �����E�H�[�Y.���[�UID);
    }

    �����E�H�[�Y.�I��(�����E�H�[�Y.�擾���� + "���̃t�@�C�����擾���܂���")
}



�����E�H�[�Y.�����ꗗ�y�[�W��� = function(url){
    var document   = IE�ړ�(url);
    var a          = document.querySelectorAll("a");
    var ��͌���   = [];

    for(var i = 0; i < a.length; i++){
       if(a[i].textContent === '\u898B\u308B'){ //����
            var ����ID   = String(a[i].onclick).match(/'(.+?)'/)[1];
            var �t�@�C�� = �����E�H�[�Y.����ID���t�@�C�����ɕϊ�(����ID)
            if(�����E�H�[�Y.�_�E�����[�h�ς݊����ꗗ.indexOf(�t�@�C��) !== -1){
                break;
            }
            ��͌���.push(����ID);
        }
        else if(a[i].textContent === '\u6B21'){ //��
            ��͌���.���̃y�[�W = a[i].href;
            break;
        }
    }

    for(var i = 0; i < ��͌���.length; i++){
        var �����y�[�W��͌��� = �����E�H�[�Y.�����y�[�W���(��͌���[i]);
        if(�����y�[�W��͌���.����){
            �����E�H�[�Y.�����t�@�C���ۑ�(�����y�[�W��͌���);
            �����E�H�[�Y.�擾����++;
        }
    }
    if(��͌���.���̃y�[�W){
        �����E�H�[�Y.�����ꗗ�y�[�W���(��͌���.���̃y�[�W);
    }
};



�����E�H�[�Y.�����y�[�W��� = function(����ID){
    var html   = HTTP_GET("https://kif-pona.heroz.jp/games/" + ����ID);
    var �\�[�X = html.substr(html.indexOf('gamedata'));

    return {
        ����ID  : ����ID,
        ��薼�O: ����ID.split(/-/)[0],
        ��薼�O: ����ID.split(/-/)[1],
        ���i��: �\�[�X.match(/dan0: "(.+?)"/)[1],
        ���i��: �\�[�X.match(/dan1: "(.+?)"/)[1],
        ����    : �\�[�X.match(/receiveMove\("(.+?)"/)[1]
    };
};



�����E�H�[�Y.�����t�@�C���ۑ� = function(��͌���){
    var kif = "";
    kif += "�J�n�����F" + �����E�H�[�Y.����ID�����Ԃɕϊ�(��͌���.����ID) + "\r\n";
    kif += "���F" + ��͌���.��薼�O + " "+ ��͌���.���i�� + "\r\n";
    kif += "���F" + ��͌���.��薼�O + " "+ ��͌���.���i�� + "\r\n";
    kif += "����F�����E�H�[�Y " + �����E�H�[�Y.���݂̃��[�h + "\r\n";
    kif += "�ꏊ�Fhttps://kif-pona.heroz.jp/games/" + ��͌���.����ID + "\r\n";
    kif +=  �����E�H�[�Y.KIF�ϊ�(��͌���.����);

    var �p�X = �����E�H�[�Y.���݂̃��[�h + '/' + �����E�H�[�Y.����ID���t�@�C�����ɕϊ�(��͌���.����ID) + '.kif';
    �t�@�C���ۑ�(�p�X, kif, "Shift_JIS");
};



�����E�H�[�Y.KIF�ϊ� = function(csa){
    var result = "";
    var ��ϊ� = {FU:'��', KY:'��' ,KE:'�j' ,GI:'��' ,KI:'��', KA:'�p', HI:'��', OU:'��', TO:'��', NY:'����' ,NK:'���j', NG:'����', UM:'�n', RY:'��'};
    var �t�ϊ� = {TO:'��', NY:'��' ,NK:'�j', NG:'��', UM:'�p', RY:'��'};
    var ����� = ['TO', 'NY', 'NK', 'NG', 'UM', 'RY'];
    var �S���� = {1:'�P', 2:'�Q', 3:'�R', 4:'�S', 5:'�T', 6:'�U', 7:'�V', 8:'�W', 9:'�X'};
    var ������ = {1:'��', 2:'��', 3:'�O', 4:'�l', 5:'��', 6:'�Z', 7:'��', 8:'��', 9:'��'};
    var �I��   = {CHECKMATE:'�l��', TORYO:'����', DISCONNECT:'��������', TIMEOUT:'�؂ꕉ��', OUTE_SENNICHI:'��������', SENNICHI:'�����'};
    var ���ԕ\ = {'10��':600, '3��':180, '10�b':3600};

    var ���c�莞�� = ���ԕ\[�����E�H�[�Y.���݂̃��[�h];
    var ���c�莞�� = ���ԕ\[�����E�H�[�Y.���݂̃��[�h];
    var ���ݐώ��� = 0;
    var ���ݐώ��� = 0;

    csa = csa.split("\t");
    for(var i = 0; i < csa.length; i++){
        var �萔   = i + 1;
        var �w���� = csa[i].split(",")[0];
        var ����   = csa[i].split(",")[1];

        if(�w����.match(/^[\+\-]/)){
            var kif  = "";
            var ��� = �w����.substr(0, 1);
            var �OX  = �w����.substr(1, 1);
            var �OY  = �w����.substr(2, 1);
            var ��X  = �w����.substr(3, 1);
            var ��Y  = �w����.substr(4, 1);
            var ��   = �w����.substr(5, 2);
            
            kif += �S����[��X]; //"���@��"�݂����ȕ\���ɑΉ����Ă��Ȃ�
            kif += ������[��Y];

            if(�����.indexOf(��) > -1 && �����E�H�[�Y.KIF�ϊ�.���蔻��(csa, i, ���, ��, �OX, �OY)){
                kif += �t�ϊ�[��] + "��";
            }
            else{
                kif += ��ϊ�[��];
            }
            kif += (�OX == 0) ? '��' : ('(' + �OX + �OY + ')');

            //����
            var ���� = Number(����.substr(1));
            if(��� === "+"){
                var �����  = ���c�莞�� - ����;
                ���ݐώ��� += �����;
                ���c�莞��  = ����;
                kif          += "  " + �����E�H�[�Y.KIF�ϊ�.����(�����, ���ݐώ���);
            }
            else{
                var �����  = ���c�莞�� - ����;
                ���ݐώ��� += �����;
                ���c�莞��  = ����;
                kif          += "  " + �����E�H�[�Y.KIF�ϊ�.����(�����, ���ݐώ���);
            }
            result += �萔 + " " + kif + "\r\n";
        }
        else{
            var �I�Ǖ����� = �w����.match(/(_WIN|DRAW)_(\w+)/);
            if(�I�Ǖ�����[2] in �I��){
                result += �萔 + " " + �I��[�I�Ǖ�����[2]] + "\r\n";
            }
            break;
        }
    }

    return result;
};



�����E�H�[�Y.KIF�ϊ�.���蔻�� = function (csa, �萔, ���, �����, �OX, �OY){
    //�O��̈ʒu�������ł����false (����ȊO��true)
    //�����Ώۂ͌��݂̎萔�܂ŁB�����̎�Ԃ̂�
    var ���� = new RegExp('^\\' + ��� + '\\d\\d' + �OX + �OY);

    for(var i = �萔; i >= 0; i -= 2){
        if(csa[i].match(����)){
            return (csa[i].match(�����)) ? false : true;
        }
    }
    return true;
};



�����E�H�[�Y.KIF�ϊ�.���� = function (�����, �ݐώ���){
    var ��� = Math.floor(����� / 60);
    var ����b = ����� % 60;
    var �ݐώ� = Math.floor(�ݐώ��� / 3600);
    var �ݐϕ� = Math.floor(�ݐώ��� % 3600 / 60);
    var �ݐϕb = �ݐώ��� % 60;

    ����b = ('0' + ����b).slice(-2);
    �ݐώ� = ('0' + �ݐώ�).slice(-2);
    �ݐϕ� = ('0' + �ݐϕ�).slice(-2);
    �ݐϕb = ('0' + �ݐϕb).slice(-2);

    return "( " + ��� + ":" + ����b + "/" + �ݐώ� + ":" + �ݐϕ� + ":" + �ݐϕb + ")";
};


�����E�H�[�Y.����ID���t�@�C�����ɕϊ� = function (����ID){
    var id = ����ID.split('-');
    return id[2] + "-" + id[0] + "-" + id[1];
}



�����E�H�[�Y.����ID�����Ԃɕϊ� = function (����ID){
    var id = ����ID.split('-');

    var �N = id[2].substr(0, 4);
    var �� = id[2].substr(4, 2);
    var �� = id[2].substr(6, 2);
    var �� = id[2].substr(9, 2);
    var �� = id[2].substr(11, 2);
    var �b = id[2].substr(13, 2);

    return �N + "/" + �� + "/" + �� + " " + �� + ":" + �� + ":" + �b;
};



�����E�H�[�Y.�I�� = function(str){
    IE.Quit();
    WScript.Echo(str);
    WScript.Quit();
};



function IE�ړ�(url){ /* �G���[���̑΍􂪕�����Ȃ� */
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



function �t�@�C���擾(file, encode){
    var stream = new ActiveXObject('ADODB.Stream');
    stream.charset = encode || 'utf-8';
    stream.Open();
    stream.loadFromFile(file);
    var contents = stream.ReadText();
    stream.close();

    return contents;
}



function �t�@�C���ۑ�(file, contents, encode){
    var stream = new ActiveXObject('ADODB.Stream');
    stream.charset = encode || 'utf-8';
    stream.Open();
    stream.WriteText(contents);
    stream.SaveToFile(file, 1);
    stream.Close();
}



function �t�@�C���ꗗ(dir, type){
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



function ���d�N���h�~(){ // https://gist.github.com/ka-ka-xyz/2718628
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
        WScript.Echo(WScript.ScriptName + "�͊��ɋN������Ă��܂��B");
        WScript.Quit(0);
    }
    else if(typeof pid === "undefined"){
        WScript.Echo("���d�N������Ɏ��s���܂����B");
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
�����E�H�[�Y();


// �E�H�[�Y�d�l�̎Q�l�y�[�W https://github.com/tosh1ki/shogiwars
// CSA�`���̎Q�l�y�[�W http://www2.computer-shogi.org/protocol/record_v22.html
// KIF�`���̎Q�l�y�[�W http://kakinoki.o.oo7.jp/kif_format.html