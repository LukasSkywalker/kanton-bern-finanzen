/*
 kanton-bern-finanzen https://github.com/KeeTraxx/kanton-bern-finanzen
 Copyright (C) 2014  Khôi Tran

 This program is free software: you can redistribute it and/or modify
 it under the terms of the GNU General Public License as published by
 the Free Software Foundation, either version 3 of the License, or
 (at your option) any later version.

 This program is distributed in the hope that it will be useful,
 but WITHOUT ANY WARRANTY; without even the implied warranty of
 MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
 GNU General Public License for more details.

 You should have received a copy of the GNU General Public License
 along with this program.  If not, see <http://www.gnu.org/licenses/>.
 */

var path = require('path');
var util = require('util');
var fs = require('fs');
var kexcel = require('kexcel');

var _ = require('underscore');

kexcel.open(path.join('brig.xlsx'), function (err, workbook) {
    var nodeData = [];
    try {
        var sheet = workbook.sheets[0];
        console.log
        var row = 2;
        while (sheet.getCellValue(row, 1) != undefined) {
            // executed for each item
            var subCategoryCode = sheet.getCellValue(row, 1);
            var categoryCode = parseInt(subCategoryCode.split('.')[0]) - 1;
            subCategoryCode = categoryCode + subCategoryCode.split('.')[1];
            var category = sheet.getCellValue(row, 2);
            var subCategory = sheet.getCellValue(row, 3);
            if (nodeData[categoryCode] == undefined) {
                nodeData[categoryCode] = {
                    code: categoryCode + '',
                    values: {},
                    name: category,
                    children: []
                };
            }
            var subData = {
                code: subCategoryCode,
                name: subCategory,
                values: {}
            };
            var column = 6;
            while (sheet.getCellValue(1, column) != undefined) {
                // executed for each year
                var header = sheet.getCellValue(1, column);
                var valid = header.indexOf('Rechnung') == 0 && header.indexOf('Aufwand') > 0;
                if (!valid) {
                    throw "Column Header is invalid: " + header;
                }
                var year = header.replace('Rechnung ', '').replace(' Aufwand', '')
                var value = sheet.getCellValue(row, column);
                var amount = parseInt(value)/1000;
                if (isNaN(amount)) {
                    amount = 0;
                }
                subData.values[year] = amount;
                if (nodeData[categoryCode].values[year] == undefined) {
                    nodeData[categoryCode].values[year] = 0;
                }
                nodeData[categoryCode].values[year] += amount; 
				column += 4;
            }
            nodeData[categoryCode].children.push(subData);
            row++;
        }
        data = [];
        for (var node in nodeData) {
            data.push(nodeData[node]);
        }
        nodeData = {
            children: data
        };
    } catch (e) {
        console.log(e);
    }

    /*fs.writeFile(path.join('..','data', 'data.json'), JSON.stringify(data, null, 4), function(){
        console.log('done!');
    });*/

    fs.writeFile(path.join('..', 'data', 'data.json'), JSON.stringify(nodeData, null, 4), function () {
        console.log('done!');
    });
});