const TelegramBot = require('node-telegram-bot-api');
const token = '432343644:AAE4h68JGzDOvWIznBmJff-zoPtxz7clxig';
var ibot,fname,fname,lname,uname,lcode,ctype,text,now,id,outmsg;
const bot = new TelegramBot(token, {polling: true});



bot.on('message', (msg) => { 
    function regpar(){
    var moment = require('moment');
    moment().format();
      id = msg.chat.id;
      ibot = msg.is_bot;
      fname = msg.from.first_name;
      lname = msg.from.last_name;
      uname = msg.from.username;
      lcode = msg.from.language_code;
      ctype = msg.chat.type;
      text = msg.text; 
      now = '['+moment().get('hour')+':'+moment().minute()+':'+moment().get('second')+']';
} 
    regpar(); 
    function getclass(day, cs){
     function dn(a){
         
         if(a != null){
             return a;
         }else{
             return '';
         }
     }
    var Excel = require('exceljs');
    var workbook = new Excel.Workbook();
    var shclass = [
[
['A1','B3','B4','B5','B6','B7','B8','B9','B10','B11'],//ПН
['A12','B14','B15','B16','B17','B18','B19','B20','B21','B22'],//BТ
['A23','B25','B26','B27','B28','B29','B30','B31','B32','B33'],//CP
['A34','B36','B37','B38','B39','B40','B41','B42','B43','B44'],//4T
['A45','B47','B48','B49','B50','B51','B52','B53','B54','B55'],//5T
['A56','B58','B59','B60','B61','B62','B63','B64','B65','B66'],//Cy
],//1
[
['A1','C3','C4','C5','C6','C7','C8','C9','C10','C11'],//ПН
['A12','C14','C15','C16','C17','C18','C19','C20','C21','C22'],//BТ
['A23','C25','C26','C27','C28','C29','C30','C31','C32','C33'],//CP
['A34','C36','C37','C38','C39','C40','C41','C42','C43','C44'],//4T
['A45','C47','C48','C49','C50','C51','C52','C53','C54','C55'],//5T
['A56','C58','C59','C60','C61','C62','C63','C64','C65','C66'],//Cy
],//2
[
['A1','D3','D4','D5','D6','D7','D8','D9','D10','D11'],//ПН
['A12','D14','D15','D16','D17','D18','D19','D20','D21','D22'],//BТ
['A23','D25','D26','D27','D28','D29','D30','D31','D32','D33'],//CP
['A34','D36','D37','D38','D39','D40','D41','D42','D43','D44'],//4T
['A45','D47','D48','D49','D50','D51','D52','D53','D54','D55'],//5T
['A56','D58','D59','D60','D61','D62','D63','D64','D65','D66'],//Cy
],//3
[
['A1','E3','E4','E5','E6','E7','E8','E9','E10','E11'],//ПН
['A12','E14','E15','E16','E17','E18','E19','E20','E21','E22'],//BТ
['A23','E25','E26','E27','E28','E29','E30','E31','E32','E33'],//CP
['A34','E36','E37','E38','E39','E40','E41','E42','E43','E44'],//4T
['A45','E47','E48','E49','E50','E51','E52','E53','E54','E55'],//5T
['A56','E58','E59','E60','E61','E62','E63','E64','E65','E66'],//Cy
],//4
[
['A1','F3','F4','F5','F6','F7','F8','F9','F10','F11'],//ПН
['A12','F14','F15','F16','F17','F18','F19','F20','F21','F22'],//BТ
['A23','F25','F26','F27','F28','F29','F30','F31','F32','F33'],//CP
['A34','F36','F37','F38','F39','F40','F41','F42','F43','F44'],//4T
['A45','F47','F48','F49','F50','F51','F52','F53','F54','F55'],//5T
['A56','F58','F59','F60','F61','F62','F63','F64','F65','F66'],//Cy
],//5
[
['A1','G3','G4','G5','G6','G7','G8','G9','G10','G11'],//ПН
['A12','G14','G15','G16','G17','G18','G19','G20','G21','G22'],//BТ
['A23','G25','G26','G27','G28','G29','G30','G31','G32','G33'],//CP
['A34','G36','G37','G38','G39','G40','G41','G42','G43','G44'],//4T
['A45','G47','G48','G49','G50','G51','G52','G53','G54','G55'],//5T
['A56','G58','G59','G60','G61','G62','G63','G64','G65','G66'],//Cy
],//6
[
['A1','H3','H4','H5','H6','H7','H8','H9','H10','H11'],//ПН
['A12','H14','H15','H16','H17','H18','H19','H20','H21','H22'],//BТ
['A23','H25','H26','H27','H28','H29','H30','H31','H32','H33'],//CP
['A34','H36','H37','H38','H39','H40','H41','H42','H43','H44'],//4T
['A45','H47','H48','H49','H50','H51','H52','H53','H54','H55'],//5T
['A56','H58','H59','H60','H61','H62','H63','H64','H65','H66'],//Cy
],//7
[
['A1','I3','I4','I5','I6','I7','I8','I9','I10','I11'],//ПН
['A12','I14','I15','I16','I17','I18','I19','I20','I21','I22'],//BТ
['A23','I25','I26','I27','I28','I29','I30','I31','I32','I33'],//CP
['A34','I36','I37','I38','I39','I40','I41','I42','I43','I44'],//4T
['A45','I47','I48','I49','I50','I51','I52','I53','I54'
,'I55'],//5T
['A56','I58','I59','I60','I61','I62','I63','I64','I65','I66'],//Cy
],//8
[
['A1','J3','J4','J5','J6','J7','J8','J9','J10','J11'],//ПН
['A12','J14','J15','J16','J17','J18','J19','J20','J21','J22'],//BТ
['A23','J25','J26','J27','J28','J29','J30','J31','J32','J33'],//CP
['A34','J36','J37','J38','J39','J40','J41','J42','J43','J44'],//4T
['A45','J47','J48','J49','J50','J51','J52','J53','J54','J55'],//5T
['A56','J58','J59','J60','J61','J62','J63','J64','J65','J66'],//Cy
],//9 (зашкварный)
[
['A1','K3','K4','K5','K6','K7','K8','K9','K10','K11'],//ПН
['A12','K14','K15','K16','K17','K18','K19','K20','K21','K22'],//BТ
['A23','K25','K26','K27','K28','K29','K30','K31','K32','K33'],//CP
['A34','K36','K37','K38','K39','K40','K41','K42','K43','K44'],//4T
['A45','K47','K48','K49','K50','K51','K52','K53','K54','K55'],//5T
['A56','K58','K59','K60','K61','K62','K63','K64','K65','K66'],//Cy
],//10 (TOP)
[
['A1','L3','L4','L5','L6','L7','L8','L9','L10','L11'],//ПН
['A12','L14','L15','L16','L17','L18','L19','L20','L21','L22'],//BТ
['A23','L25','L26','L27','L28','L29','L30','L31','L32','L33'],//CP
['A34','L36','L37','L38','L39','L40','L41','L42','L43','L44'],//4T
['A45','L47','L48','L49','L50','L51','L52','L53','L54','L55'],//5T
['A56','L58','L59','L60','L61','L62','L63','L64','L65','L66'],//Cy
],//11
]
    var day = day - 1;
    var cs = cs - 1;
    workbook.xlsx.readFile('time.xlsx')
    .then(function() {
       var ws = workbook.getWorksheet('Лист1');
         outmsg = dn(ws.getCell(shclass[cs][day][0]).value)+'\n'+dn(ws.getCell(shclass[cs][day][1]).value)+'\n'+dn(ws.getCell(shclass[cs][day][2]).value)+'\n'+dn(ws.getCell(shclass[cs][day][3]).value)+'\n'+dn(ws.getCell(shclass[cs][day][4]).value)+'\n'+dn(ws.getCell(shclass[cs][day][5]).value)+'\n'+dn(ws.getCell(shclass[cs][day][6]).value)+'\n'+dn(ws.getCell(shclass[cs][day][7]).value)+'\n'+dn(ws.getCell(shclass[cs][day][8]).value)+'\n'+dn(ws.getCell(shclass[cs][day][9]).value);
        
    });
    return outmsg;
    }
    function say(umsg,chatid){
        bot.sendMessage(id,' '+umsg);
        console.log(now+' SchoolBot: '+umsg);
    }
     console.log(now+' '+uname+': '+text);
    console.log('debug: '+getclass(1,1));//костыль
    say(getclass(2,10));
    
});
