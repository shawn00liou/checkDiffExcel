const path = require('path');
const Excel = require('exceljs');
const filesJs = require('./files.js');

(function () {
  //最新的
  const workbook = new Excel.Workbook();
  const workbookOld = new Excel.Workbook();
  const workbookFinally = new Excel.Workbook();//最終輸出

  const newJson = []; //輸出的整理
  const promise1 = new Promise((resolve, reject) => {
    workbook.xlsx.readFile('Inspection_20201016.xlsx').then(function () {
      const worksheet = workbook.getWorksheet('MySheet');

      worksheet.eachRow({ includeEmpty: true }, function (row, rowNumber) {
        /** 一列列讀出來 */
        if (rowNumber > 1) {
          const currRow = worksheet.getRow(rowNumber);
          const RowData = currRow._cells.map((item, ind) => {
            return clearFormat(currRow.getCell(ind + 1).value);
          });
          newJson.push(RowData);
        }
      });

      resolve();
    }, errorHandler);
  });

  const oldJson = [];
  const promise2 = new Promise((resolve, reject) => {
    workbookOld.xlsx.readFile('Inspection_FriSep112020.xlsx').then(function () {
      const worksheet = workbookOld.getWorksheet('MySheet');

      worksheet.eachRow({ includeEmpty: true }, function (row, rowNumber) {
        /** 一列列讀出來 */
        if (rowNumber > 1) {
          const currRow = worksheet.getRow(rowNumber);
          const RowData = currRow._cells.map((item, ind) => {
            return clearFormat(currRow.getCell(ind + 1).value);
          });

          oldJson.push(RowData);
        }
      });

      resolve();
    }, errorHandler);
  });

  //產出有合併欄位的 Excels
  /**
   * 繁體中文	zh-tw
   * 簡體中文	zh-cn
   * 英文	en
   * 越文	vi
   * 泰文	th
   * 馬來文	ms
   * 印尼文	id
   * 印度文	hi
   */
  const headerLangKey = {
    key: 'key',
    'zh-cn': '简体',
    'zh-tw': '繁體',
    en: '英文',
    th: '泰文',
    vi: '越文',
    hi:'印地文',
    rowid: 'rowid',
  };

  const langEnum = [
    "key",
    "zh-cn",
    "zh-tw",
    "en",
    "th",
    "vi",
    "hi",
    "rowid"
  ]

  Promise.all([promise1, promise2]).then(() => {

    const finallyRow = newJson.map((item,index)=>{
      // console.log(item[0],index,'//',oldJson.length)
      let check = false;
      oldJson.forEach((it,key)=>{
        if(item[0]===it[0]){
         check = (item[1]===it[1] && item[2]===it[2] && item[3]===it[3])?false:true;
        //  if(check){
        //    console.log(item[0],item[3],'!=',it[3])
        //  }
          return;
        }
      })

      if(check || !item[item.length-2]){
        const rowobj = {}
        langEnum.forEach((langval,langindex)=>{
          rowobj[langval] = item[langindex];
        });

        return rowobj
      }
    });
    console.log(finallyRow.filter((it) => it).length,'/////',newJson.length);
    const worksheetFinally = workbookFinally.addWorksheet('MySheet');
    const excelColumn = Object.keys(headerLangKey).map((it) => {
      return { header: headerLangKey[it], key: it, width: 100 };
    });
    worksheetFinally.columns = excelColumn;
    worksheetFinally.addRows(finallyRow.filter((it) => it));
    (async function () {
      return await workbookFinally.xlsx.writeFile('Inspection_20201017.xlsx').then(async () => {
        console.log('success!!!!!')
      }, errorHandler);
    })();

  });

  console.log('INI');
})();

function errorHandler(err) {
  if (err) {
    console.log(err);
    throw err;
  }
}

function clearFormat(params) {
  if (typeof params === 'object' && params && params.richText) {
    const ar = Object.values(params.richText).map((item) => {
      if (item.text) {
        return item.text;
      }
    });
    return ar.join('');
  }
  return params;
}
