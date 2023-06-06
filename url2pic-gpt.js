const ExcelJS = require("exceljs");
const https = require("https");
const fs = require("fs");

// 读取 Excel 文件
const workbook = new ExcelJS.Workbook();
workbook.xlsx
  .readFile("a.xlsx")
  .then(() => {
    const worksheet = workbook.getWorksheet(1);

    // 遍历每个单元格
    worksheet.eachRow((row) => {
      row.eachCell((cell) => {
        const cellValue = cell.value;

        // 判断单元格内容是否为图片链接
        if (typeof cellValue === "string" && cellValue.startsWith("http")) {
          // 下载图片并替换单元格内容
          downloadImage(cellValue, (imagePath) => {
            cell.value = {
              hyperlink: imagePath,
              text: "点击查看",
            };
            cell.style = { hyperlink: true };

            // 保存修改后的 Excel 文件
            workbook.xlsx
              .writeFile("b.xlsx")
              .then(() => {
                console.log("转换完成");
              })
              .catch((err) => {
                console.error("保存 Excel 文件失败:", err);
              });
          });
        }
      });
    });
  })
  .catch((err) => {
    console.error("读取 Excel 文件失败:", err);
  });

// 下载图片
function downloadImage(url, callback) {
  const imagePath = `images/${Date.now()}.jpg`; // 保存图片的路径，可根据需求修改

  https
    .get(url, (response) => {
      const fileStream = fs.createWriteStream(imagePath);
      response.pipe(fileStream);

      fileStream.on("finish", () => {
        fileStream.close();
        callback(imagePath);
      });
    })
    .on("error", (err) => {
      console.error("下载图片失败:", err);
    });
}
