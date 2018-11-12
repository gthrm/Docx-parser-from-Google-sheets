var fs = require('fs');
var zip = new require('node-zip')();
var path = require ('path');
var axios = require('axios');

/* читаем файлы архива в память */
var f01 = fs.readFileSync( path.resolve(__dirname,'doc/_rels/.rels') );
var f02 = fs.readFileSync( path.resolve(__dirname,'doc/docProps/app.xml') );
var f03 = fs.readFileSync( path.resolve(__dirname,'doc/docProps/core.xml') );
var f04 = fs.readFileSync( path.resolve(__dirname,'doc/word/_rels/document.xml.rels') );
var f05 = fs.readFileSync( path.resolve(__dirname,'doc/word/theme/theme1.xml') );
var f06 = fs.readFileSync( path.resolve(__dirname,'doc/word/document.xml') );
var f07 = fs.readFileSync( path.resolve(__dirname,'doc/word/fontTable.xml') );
var f08 = fs.readFileSync( path.resolve(__dirname,'doc/word/numbering.xml') );
var f09 = fs.readFileSync( path.resolve(__dirname,'doc/word/settings.xml') );
var f10 = fs.readFileSync( path.resolve(__dirname,'doc/word/styles.xml') );
// var f11 = fs.readFileSync( path.resolve(__dirname,'doc/word/stylesWithEffects.xml') );
var f12 = fs.readFileSync( path.resolve(__dirname,'doc/word/webSettings.xml') );
var f13 = fs.readFileSync( path.resolve(__dirname,'doc/[Content_Types].xml') );
/* тут все остальные файлы */

/* создаём zip-объект */
zip.file('_rels/.rels', f01);
zip.file('docProps/app.xml', f02);
zip.file('docProps/core.xml', f03);
zip.file('word/_rels/document.xml.rels', f04);
zip.file('word/theme/theme1.xml', f05);
zip.file('word/document.xml', f06);
zip.file('word/fontTable.xml', f07);
zip.file('word/numbering.xml', f08);
zip.file('word/settings.xml', f09);
zip.file('word/styles.xml', f10);
// zip.file('word/stylesWithEffects.xml', f11);
zip.file('word/webSettings.xml', f12);
zip.file('[Content_Types].xml', f13);

axios.get("https://spreadsheets.google.com/feeds/list/1Yu4qXyEXOU82OZj6WeE8kju204g1PSRGlto4cw5JEz4/od6/public/values?alt=json").then(({data})=>write_record(data['feed']['entry']));

//запрос отправить на указанный адрес, где заменить "1Yu4qXyEXOU82OZj6WeE8kju204g1PSRGlto4cw5JEz4"
function write_record(data_lol) {    
    for(i=0; i<data_lol.length; i++){
        var body = String(f06).replace("ИМЯ_ЭКСПЕРТА", ""+data_lol[i]['gsx$name']['$t']);
        //к наименованием в клетках добавить "gsx$" -> name = gsx$name

        zip.file('word/document.xml', body); //, {mode: 0777}
        var data = zip.generate({base64:false,compression:'DEFLATE'});
        fs.writeFileSync(path.resolve(__dirname,"result/"+ data_lol[i]['gsx$key']['$t'] + ".docx") , data, 'binary');
        console.log(data_lol[i]['gsx$key']['$t'] + ".docx сохранен");
    };
};