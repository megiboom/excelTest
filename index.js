var Excel = require('exceljs');
var groupJson = require('./groupTest');
var moduleData = require('./test.json')

var workbook = new Excel.Workbook();

var color=['FFC6E0B4','FFFFFF99','FFFFE699','FFC6E0B4','FFCCFFCC','FFBDD7EE']

function numToSSColumn(num){
  var s = '', t;
  var n = num;
  while (num > 0) {
    t = (num - 1) % 26;
    s = String.fromCharCode(65 + t) + s;
    num = (num - t)/26 | 0;
  }
  return s || undefined;
}

function makeExcel(data){
    workbook.xlsx.readFile('template.xlsx')
    .then(function() {
        var worksheet = workbook.getWorksheet(1);

        var rowNum = 3;
        var cellNum = 2;
        var groupLength = data.length;

        for(var j =0; j<groupLength; j++){
            
            var moduleLength = 10;
            var moduleLength = data[j].moduleList.length;

            /* 대분류 시작 */
            //시작행, 시작열, 끝행, 끝열
            worksheet.mergeCells(rowNum, cellNum, rowNum, cellNum+3*moduleLength-1);
            worksheet.getCell(rowNum,cellNum).alignment = {
                vertical: 'middle', 
                horizontal: 'center'
            }

            const randomColor=color[j%color.length];
            
            worksheet.getCell(rowNum,cellNum).fill = {
                type:'pattern',
                pattern:'solid',
                fgColor:{argb:randomColor}
                //fgColor:{argb:'CCFFEE'}
            }
            worksheet.getCell(rowNum,cellNum).border = {top: {style:'medium'},left: {style:'thin'},bottom: {style:'thin'},right: {style:'thin'}};

            var row = worksheet.getRow(rowNum);        
            row.getCell(cellNum).value='154kV S/S (154kV)'+j  //대분류 명명
            row.getCell(cellNum).font={bold:true}
            /* 대분류 끝 */

            /* 중분류 시작 */
            for(var i=0; i<moduleLength;i++){

                worksheet.mergeCells(rowNum+1, cellNum+3*i, rowNum+1, (cellNum+3*i)+2);
                worksheet.getCell(rowNum+1,cellNum+3*i).alignment = {
                    vertical: 'middle', 
                    horizontal: 'center'
                }
                worksheet.getCell(rowNum+1,cellNum+3*i).fill = {
                    type:'pattern',
                    pattern:'solid',
                    fgColor:{argb:randomColor}
                }
                worksheet.getCell(rowNum+1,cellNum+3*i).border = {top: {style:'thin'},left: {style:'thin'},bottom: {style:'thin'},right: {style:'thin'}};
            
                row = worksheet.getRow(rowNum+1);
                row.getCell(cellNum+3*i).value='(1)154kV GIS MAIN반(☞280,000)'
                row.getCell(cellNum+3*i).font={bold:true}

                /* 소분류 시작 */
                row = worksheet.getRow(rowNum+2);
                row.getCell(cellNum+3*i).value='전류[A]'
                worksheet.getCell(rowNum+2,cellNum+3*i).border = {top: {style:'thin'},left: {style:'thin'},bottom: {style:'thin'},right: {style:'hair'}};
                worksheet.getCell(rowNum+2,cellNum+3*i).fill = {
                    type:'pattern',
                    pattern:'solid',
                    fgColor:{argb:randomColor}
                }
                row.getCell(cellNum+3*i+1).value='지침'
                worksheet.getCell(rowNum+2,cellNum+3*i+1).border = {top: {style:'thin'},left: {style:'hair'},bottom: {style:'thin'},right: {style:'hair'}};
                worksheet.getCell(rowNum+2,cellNum+3*i+1).fill = {
                    type:'pattern',
                    pattern:'solid',
                    fgColor:{argb:randomColor}
                }
                row.getCell(cellNum+3*i+2).value='전력량'
                worksheet.getCell(rowNum+2,cellNum+3*i+2).border = {top: {style:'thin'},left: {style:'hair'},bottom: {style:'thin'},right: {style:'thin'}};
                worksheet.getCell(rowNum+2,cellNum+3*i+2).fill = {
                    type:'pattern',
                    pattern:'solid',
                    fgColor:{argb:randomColor}
                }

                //합계
                row = worksheet.getRow(rowNum+28);
                var colNum = numToSSColumn(cellNum+3*i+2)
                row.getCell(cellNum+3*i+2).value={ formula: "SUM("+colNum+(rowNum+3)+","+colNum+(rowNum+27)+")"}
                row.getCell(cellNum+3*i+2).numFmt = '#,##.0';
                row.getCell(cellNum+3*i+2).font={bold:true,color: { argb: 'FFFF0000' }}
                row.getCell(cellNum+3*i).border = {top: {style:'thin'},left: {style:'hair'},bottom: {style:'hair'},right: {style:'hair'}};
                row.getCell(cellNum+3*i+1).border = {top: {style:'thin'},left: {style:'hair'},bottom: {style:'hair'},right: {style:'hair'}};
                row.getCell(cellNum+3*i+2).border = {top: {style:'thin'},left: {style:'hair'},bottom: {style:'hair'},right: {style:'thin'}};
                //평균
                row = worksheet.getRow(rowNum+29);
                var colNum = numToSSColumn(cellNum+3*i+2)
                row.getCell(cellNum+3*i+2).value={ formula: "Average("+colNum+(rowNum+3)+","+colNum+(rowNum+27)+")"}
                row.getCell(cellNum+3*i+2).numFmt = '#,##.0';
                row.getCell(cellNum+3*i+2).font={bold:true,color: { argb: 'FF800080' }}
                row.getCell(cellNum+3*i).border = {top: {style:'thin'},left: {style:'hair'},bottom: {style:'hair'},right: {style:'hair'}};
                row.getCell(cellNum+3*i+1).border = {top: {style:'thin'},left: {style:'hair'},bottom: {style:'hair'},right: {style:'hair'}};
                row.getCell(cellNum+3*i+2).border = {top: {style:'thin'},left: {style:'hair'},bottom: {style:'hair'},right: {style:'thin'}};
                //최대전력
                row = worksheet.getRow(rowNum+30);
                var colNum = numToSSColumn(cellNum+3*i+2)
                row.getCell(cellNum+3*i+2).value={ formula: "MAX("+colNum+(rowNum+3)+","+colNum+(rowNum+27)+")"}
                row.getCell(cellNum+3*i+2).numFmt = '#,##.0';
                row.getCell(cellNum+3*i+2).font={bold:true,color: { argb: 'FF0000FF' }}
                row.getCell(cellNum+3*i).border = {top: {style:'thin'},left: {style:'hair'},bottom: {style:'hair'},right: {style:'hair'}};
                row.getCell(cellNum+3*i+1).border = {top: {style:'thin'},left: {style:'hair'},bottom: {style:'hair'},right: {style:'hair'}};
                row.getCell(cellNum+3*i+2).border = {top: {style:'thin'},left: {style:'hair'},bottom: {style:'hair'},right: {style:'thin'}};
                // 부하량
                row = worksheet.getRow(rowNum+31);
                row.getCell(cellNum+3*i+2).font={bold:true,color: { argb: 'FFFF00FF' }}
                row.getCell(cellNum+3*i).border = {top: {style:'thin'},left: {style:'hair'},bottom: {style:'medium'},right: {style:'hair'}};
                row.getCell(cellNum+3*i+1).border = {top: {style:'thin'},left: {style:'hair'},bottom: {style:'medium'},right: {style:'hair'}};
                row.getCell(cellNum+3*i+2).border = {top: {style:'thin'},left: {style:'hair'},bottom: {style:'medium'},right: {style:'thin'}};
                /* 소분류 끝 */


                var rows = worksheet.getRows(rowNum+3,rowNum+22);
            
                rows.forEach((value) => {
                    value.getCell(cellNum+3*i).value = 100000+j; //전류
                    value.getCell(cellNum+3*i).border = {top: {style:'hair'},left: {style:'hair'},bottom: {style:'hair'},right: {style:'hair'}};
                    value.getCell(cellNum+3*i).numFmt='#,##'         
                    //지침
                    value.getCell(cellNum+3*i+1).border = {top: {style:'hair'},left: {style:'hair'},bottom: {style:'hair'},right: {style:'hair'}};

                    value.getCell(cellNum+3*i+2).value = 100000+i; //전력량
                    value.getCell(cellNum+3*i+2).border = {top: {style:'hair'},left: {style:'hair'},bottom: {style:'hair'},right: {style:'thin'}};
                    value.getCell(cellNum+3*i+2).numFmt='#,##.0'
                    value.commit();
                })
            }
            row.commit();
            /* 중분류 끝 */
            cellNum=cellNum+(3*moduleLength);
        }

        return "complete";
        //return workbook.xlsx.writeFile('new.xlsx');
    }).then(()=>{
        workbook.xlsx.writeFile('new.xlsx');
    })
}

makeExcel(groupJson.groupBy(moduleData));