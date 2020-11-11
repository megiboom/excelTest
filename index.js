var Excel = require('exceljs');
var workbook = new Excel.Workbook();

workbook.xlsx.readFile('template.xlsx')
    .then(function() {
        var worksheet = workbook.getWorksheet(1);

        var rowNum = 2;
        var cellNum = 3;
        var moduleLength = 3;

        /* 대분류 시작 */
        //시작행, 시작열, 끝행, 끝열
        worksheet.mergeCells(rowNum, cellNum, rowNum, cellNum+3*moduleLength-1);
        worksheet.getCell(rowNum,cellNum).alignment = {
            vertical: 'middle', 
            horizontal: 'center'
        }
        worksheet.getCell(rowNum,cellNum).fill = {
            type:'pattern',
            pattern:'solid',
            fgColor:{argb:'CCFFEE'}
        }
        worksheet.getCell(rowNum,cellNum).border = {top: {style:'medium'},left: {style:'thin'},bottom: {style:'thin'},right: {style:'thin'}};

        var row = worksheet.getRow(rowNum);        
        row.getCell(cellNum).value='154kV S/S (154kV)'
        /* 대분류 끝 */

        /* 중분류 시작 */
        for(var i=0; i<moduleLength;i++){//0 (3,3,2,5) 1 (3,6,2,8) 2 (3,9,2,11) 3>5>8>11
            worksheet.mergeCells(rowNum+1, cellNum+3*i, rowNum+1, (cellNum+3*i)+2);
            worksheet.getCell(rowNum+1,cellNum+3*i).alignment = {
                vertical: 'middle', 
                horizontal: 'center'
            }
            worksheet.getCell(rowNum+1,cellNum+3*i).fill = {
                type:'pattern',
                pattern:'solid',
                fgColor:{argb:'CCFFEE'}
            }
            worksheet.getCell(rowNum+1,cellNum+3*i).border = {top: {style:'thin'},left: {style:'thin'},bottom: {style:'thin'},right: {style:'thin'}};
    
            row = worksheet.getRow(rowNum+1);
            row.getCell(cellNum+3*i).value='(1)154kV GIS MAIN반(☞280,000)'

            row = worksheet.getRow(rowNum+2);
            row.getCell(cellNum+3*i).value='전류'
            row.getCell(cellNum+3*i+1).value='지침'
            row.getCell(cellNum+3*i+2).value='전력량'
        }
        /* 중분류 끝 */


        row.commit();
        return workbook.xlsx.writeFile('new.xlsx');
    })