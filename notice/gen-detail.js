const ExcelJS = require('exceljs');
const wb = new ExcelJS.Workbook();
const fileName = '注意事項清單.xlsx';


wb.xlsx.readFile(fileName).then(() => {
    const ws = wb.getWorksheet('注意事項');
    let masterId = 0;
    let tempData = null;

    // feature Id index
    let featureIdIndex = 2;
    // tag index
    let tagIndex = 3
    // sort index
    let sortIndex = 4;
    // label index
    let labelIndex = 5;
    // content index
    let contentIndex = 6;
    // english content index
    let enContentIndex = 7;
    // color code index
    let colorCodeIndex = 8;
    // bolder index
    let bolderIndex = 9;

    let featureId = '';
    let tag = '';
    let sort = 0;
    let label = '';
    let content = '';
    let enContent = '';
    let colorCode = '';
    let bolder = 0;

    let sql = "INSERT INTO BNKMFeatureNoticeDetail VALUES";
    const sql_item = [];

    for (let row = 2; row <= ws.actualRowCount; row++) {
        for (let col = featureIdIndex; col <= bolderIndex; col++) {
            const cellValue = ws.getRow(row).getCell(col).text
            switch (col) {
                case featureIdIndex:
                    featureId = cellValue;
                    break;
                case tagIndex:
                    tag = cellValue;
                    break;
                case sortIndex:
                    sort = Number(cellValue);
                    break;
                case labelIndex:
                    label = cellValue ? cellValue : '';
                    break;
                case contentIndex:
                    content = cellValue ? cellValue : '';
                    content = content.trim();
                    break;
                case enContentIndex:                               
                    enContent = cellValue ? cellValue : '';
                    enContent = enContent.trim();   
                    break;
                case colorCodeIndex:
                    colorCode = cellValue ? cellValue : '';
                    break;
                case bolderIndex:
                    bolder = Number(cellValue) || (cellValue !== null && cellValue !== undefined && cellValue !== "")
                        ? 1
                        : 0;
                    if (tempData && featureId == tempData.featureId && tag == tempData.tag) {
                    } else {
                        tempData = { featureId, tag };
                        masterId++;
                    }

                    if (enContent.includes('\'')) {
                        enContent = enContent.replace(/'/g, '\'\'');
                    }
             
                    sql_item.push(
                        ` (${masterId}, '${label}', N'${content}', '${enContent}', ${sort}, ${bolder}, '${colorCode}', GETDATE(), GETDATE()) `
                    );
                    break;
                default:
                    featureId = '';
                    tag = '';
                    sort = 0;
                    label = '';
                    content = '';
                    enContent = '';
                    colorCode = '';
                    bolder = 0;
                    break;
            }           
        } 
    }

    if (sql_item.length) {
        console.log(`共 ${sql_item.length} 項 Detail`);

        sql += `${sql_item.join(",")}`;

        const fs = require("fs");

        fs.writeFile("./notice/detail.txt", sql, (err) => {
            if (err) {
                console.error(err);
                return;
            }
            //文件寫入成功。
        });
    } else {
        console.log(`Empty Data`);
    }
});
