import { example } from "../../example";
import * as Papa from "papaparse";
import moment from "moment";
const queryURL = "https://dummyjson.com/products";
const refreshCheckbox = $("#refreshCheckbox");
const refreshTime = $("#refreshTime");
const fetchDataBtn = $("#fetchDataBtn");
const fetchDataBtnTxt = $("#fetchDataBtnTxt");
const fetchDataBtnLoader = $("#fetchDataBtnLoader").hide();
let interval;
let minutes = 1;
Office.onReady(() => {
  // fetchResponse();
});
fetchDataBtn.on("click", function () {
  btnStyle(true);
  fetchResponse();
});
refreshTime.on("change", function () {
  stopRefresh();
  startRefresh();
});
refreshCheckbox.on("click", function () {
  if (refreshCheckbox.prop("checked")) {
    refreshTime.prop("disabled", false);
    stopRefresh();
    startRefresh();
  } else {
    refreshTime.prop("disabled", true);
    stopRefresh();
  }
});
function fetchResponse() {
  $.ajax({
    url: queryURL,
    method: "get"
  }).done(function (response, status, xhr) {
    // let responseType = xhr.getResponseHeader("content-type") || "";
    // if (responseType.indexOf("json") > -1) {
    //   runInExcel(response, "JSON");
    // } else {
    //   runInExcel(response, "CSV");
    // }
    // btnStyle(false);
    runInExcel(example, "CSV");
    btnStyle(false);
  }).fail(function () {
    console.log("Error ocurred during data fetch");
    btnStyle(false);
  });
}
function startRefresh() {
  let value = Number($("#refreshTime").val());
  if (value >= 1) {
    minutes = value * 1000 * 60;
    if (!interval) {
      interval = setInterval(fetchResponse, minutes);
    }
  }
}
function stopRefresh() {
  clearInterval(interval);
  interval = null;
}
function btnStyle(value) {
  if (value) {
    fetchDataBtnTxt.text("Fetching...");
    fetchDataBtnLoader.show();
  } else {
    fetchDataBtnTxt.text("Fetch Data");
    fetchDataBtnLoader.hide();
  }
}
function resize(array, maxLength) {
  for (let i = 0; i < array.length; ++i) {
    array[i] = array[i].concat(new Array(maxLength - array[i].length));
  }
}
function convertDate(array) {
  let regex = new RegExp(/[0-9]{4}-[0-9]{2}-[0-9]{2}T[0-9]{2}:[0-9]{2}:[0-9]{2}(\.[0-9]+)?([Zz]|([\+-])([01]\d|2[0-3]):?([0-5]\d)?)?/);
  for (let i = 0; i < array.length; ++i) {
    array[i] = array[i].map(obj => {
      if (regex.test(obj)) {
        return moment(obj).format("ddd, hA");
      } else {
        return obj;
      }
    });
  }
  return array;
}
function maxArrayLength(array) {
  let max = 0;
  for (const row of array) {
    if (row.length > max) {
      max = row.length;
    }
  }
  return max;
}
async function runInExcel(response, type) {
  try {
    await Excel.run(async context => {
      let maxLength;
      const workSheet = context.workbook.worksheets.getActiveWorksheet();
      switch (type) {
        case "JSON":
          let jsonData = [];
          let headers = ["Brand", "Item", "Price", "Rating", "Stock"];
          response.products.forEach(arr => {
            jsonData.push([arr.brand, arr.title, arr.price, arr.rating, arr.stock]);
          });
          jsonData.unshift(headers);
          maxLength = maxArrayLength(jsonData);
          const range = workSheet.getRangeByIndexes(0, 0, jsonData.length, maxLength);
          range.values = jsonData;
          range.untrack();
          break;
        case "CSV":
          let parsedCsvData = Papa.parse(response);
          let csvData = convertDate(parsedCsvData.data);
          maxLength = maxArrayLength(csvData);
          resize(csvData, maxLength);
          if (csvData.length > 0 && maxLength > 0) {
            const range = workSheet.getRangeByIndexes(0, 0, csvData.length, maxLength);
            range.values = csvData;
            range.untrack();
          }
          break;
      }
      const color = workSheet.getUsedRange().getRow(0);
      color.format.fill.color = "#b2b4cd";
      workSheet.getUsedRange().format.autofitColumns();
      workSheet.getUsedRange().format.autofitRows();
      workSheet.activate();
      await context.sync();
    });
  } catch (error) {
    console.error(error);
  }
}