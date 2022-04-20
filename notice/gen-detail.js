const readXlsxFile = require("read-excel-file/node");
const fs = require("fs");

readXlsxFile(fs.createReadStream("./注意事項清單.xlsx")).then((rows) => {
  if (!rows.length) {
    return;
  }

  let count = 0;

  let masterId = 0;
  let tempData = null;

  // feature Id index
  let featureIdIndex = -1;
  // tag index
  let tagIndex = -1;
  // label index
  let labelIndex = -1;
  // sort index
  let sortIndex = -1;
  // content index
  let contentIndex = -1;
  // english content index
  let enContentIndex = -1;
  // color code index
  let colorCodeIndex = -1;
  // bolder index
  let bolderIndex = -1;

  let sql = "INSERT INTO BNKMFeatureNoticeDetail VALUES";
  const sql_item = [];

  const errorList = [];

  for (const row of rows) {
    count++;

    //#region Index

    if (count == 1) {
      //#region label index

      const labelIndex_search = row.indexOf("標號");
      if (labelIndex_search >= 0) {
        labelIndex = labelIndex_search;
      }

      //#endregion

      //#region sort index

      const sort_search = row.indexOf("行數");
      if (sort_search >= 0) {
        sortIndex = sort_search;
      }

      //#endregion

      //#region content index

      const content_search = row.indexOf("內容");
      if (content_search >= 0) {
        contentIndex = content_search;
      }

      //#endregion

      //#region english content index

      const enContent_search = row.indexOf("English內容");
      if (enContent_search >= 0) {
        enContentIndex = enContent_search;
      }

      //#endregion

      //#region color code index

      const colorCode_search = row.indexOf("色碼");
      if (colorCode_search >= 0) {
        colorCodeIndex = colorCode_search;
      }

      //#endregion

      //#region content index

      const bolder_search = row.indexOf("粗體");
      if (bolder_search >= 0) {
        bolderIndex = bolder_search;
      }

      //#endregion

      //#region feature Id

      const featureId_find = row.indexOf("功能ID");
      if (featureId_find >= 0) {
        featureIdIndex = featureId_find;
      }

      //#endregion

      //#region tag

      const tag_find = row.indexOf(
        "頁面名稱(index、input、confirm、result.....)"
      );
      if (tag_find >= 0) {
        tagIndex = tag_find;
      }

      //#endregion

      continue;
    } else if (
      labelIndex < 0 ||
      sortIndex < 0 ||
      contentIndex < 0 ||
      enContentIndex < 0 ||
      colorCodeIndex < 0 ||
      bolderIndex < 0
    ) {
      throw new Error("欄位title不正確!");
    }

    //#endregion

    const featureId = row[featureIdIndex];
    const tag = row[tagIndex];
    const label = row[labelIndex] ?? "";
    const sort = Number(row[sortIndex]);
    const content = row[contentIndex];
    let enContent = row[enContentIndex];
    let bolder = row[bolderIndex];
    let colorCode = row[colorCodeIndex];

    //#region Check Exist

    if (!sort) {
      errorList.push(`第${count}列 sort 資料錯誤!`);
    }

    if (!content) {
      errorList.push(`第${count}列 content 資料錯誤!`);
    }

    //#endregion

    if (errorList.length) {
      throw new Error(errorList.join("\n"));
    }

    if (tempData && featureId == tempData.featureId && tag == tempData.tag) {
    } else {
      tempData = { featureId, tag };
      masterId++;
    }

    enContent = enContent ? enContent : "";
    if (enContent.includes('\'')){
      enContent = enContent.replace(/'/g, '\'\'');
    }
    
    colorCode = colorCode ? colorCode : "";
    bolder =
      bolder || (bolder !== null && bolder !== undefined && bolder !== "")
        ? 1
        : 0;

    sql_item.push(
      ` (${masterId}, '${label}', N'${content}', '${enContent}', ${sort}, ${bolder}, '${colorCode}', GETDATE(), GETDATE()) `
    );
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
