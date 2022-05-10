const readXlsxFile = require("read-excel-file/node");
const fs = require("fs");

readXlsxFile(fs.createReadStream("./注意事項清單.xlsx")).then((rows) => {
  if (!rows.length) {
    return;
  }

  // 資料筆數
  let count = 0;
  // SQL
  let sql = "INSERT INTO BNKMFeatureNotice VALUES";
  // feature Id index
  let featureIdIndex = -1;
  // tag index
  let tagIndex = -1;

  const unique_list = [];

  let tempData = null;
  // previous item is unique error
  let isPreviousUniqueError = false;

  const errorList = [];

  for (const row of rows) {
    count++;

    //#region Column title Index

    if (count == 1) {
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
    } else if (featureIdIndex < 0 || tagIndex < 0) {
      throw new Error("欄位title不正確!");
    }

    //#endregion

    const featureId = row[featureIdIndex];
    const tag = row[tagIndex];

    //#region Check Exist

    if (!featureId) {
      errorList.push(`第${count}列 'featureId' 資料錯誤!`);
      continue;
    }

    if (!tag) {
      errorList.push(`第${count}列 'tag' 資料錯誤!`);
      continue;
    }

    //#endregion

    //#region Check Duplicate

    if (
      isPreviousUniqueError &&
      featureId === tempData.featureId &&
      tag === tempData.tag
    ) {
      errorList.push(
        `第${count}列 unique error! (FeatureId: ${featureId}, Tag: ${tag})`
      );
      isPreviousUniqueError = true;
    } else if (
      tempData &&
      (featureId !== tempData.featureId || tag !== tempData.tag) &&
      unique_list.some((x) => x.featureId == featureId && x.tag == tag)
    ) {
      errorList.push(
        `第${count}列 unique error! (FeatureId: ${featureId}, Tag: ${tag})`
      );
      isPreviousUniqueError = true;
    } else {
      isPreviousUniqueError = false;
    }

    tempData = { featureId, tag };

    //#endregion

    //針對FeatureId 欄位進行@分割 取得FeatureId資訊 與 Scenario 場景欄位
    const features = featureId.split('@');
    if (unique_list.some((x) => x.featureId == features[0] && x.tag == tag && x.scenario == features[1])) {
      continue;
    } else {
      unique_list.push({ featureId:features[0], tag, scenario:features[1] });
    }
  }

  if (errorList.length) {
    throw new Error(errorList.join("\n"));
  }

  if (unique_list.length) {
    console.log(`共 ${unique_list.length} 項 Mater Key`);
    const item = unique_list.map(
      (x) => ` ('${x.featureId}', '${x.tag}', `.concat(x.scenario?`'${x.scenario}'`:'NULL', ', GETDATE(), GETDATE()) ')
    );

    sql += `${item.join(",")}`;

    const fs = require("fs");

    fs.writeFile("./notice/master.txt", sql, (err) => {
      if (err) {
        console.error(err);
        return;
      }
    });
    //文件寫入成功。
  } else {
    console.log(`Empty Data`);
  }
});
