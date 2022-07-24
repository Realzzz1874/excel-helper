// change img url to img in excel

const ExcelJS = require("exceljs");
const path = require("path");
const fs = require("fs");
const request = require("request");
const log = require("single-line-log").stdout;

const getBuffer = (img_url) => {
  return new Promise((resolve, reject) => {
    request(
      {
        url: img_url,
        encoding: null,
      },
      (error, resp, body) => {
        if (body) {
          resolve(body);
        } else {
          resolve();
        }
      }
    );
  });
};

const saveImg = async (workbook, img_obj_arr) => {
  const len = img_obj_arr.length;
  for (let i = 0; i < img_obj_arr.length; i++) {
    const { sheet, img_url, position } = img_obj_arr[i];
    const buff = await getBuffer(img_url);
    if (buff) {
      log(`正在处理第 ${i + 1} / ${len} 个图片 \n\n`);
      const img_base64 = buff.toString("base64");
      const img_id = workbook.addImage({
        base64: img_base64,
        extension: "jpeg",
      });
      const sheet_row = sheet.getRow(position.row);
      sheet_row.height = 100;

      sheet.addImage(img_id, {
        tl: { col: position.col - 0.5, row: position.row - 0.5 },
        // br: { col: 3.5, row: 5.5 },
        ext: { width: 100, height: 100 },
        hyperlinks: {
          hyperlink: img_url,
          tooltip: `${img_url}`,
        },
      });
    }
  }
};

const start = async () => {
  const fdirs = await fs.readdirSync(path.join(__dirname, "./"));
  const f = fdirs.find((f) => !f.startsWith(".") && f.endsWith(".xlsx"));
  await console.info(`\n找到 excel 文件：${f}\n`);
  const xlsx_file = path.join(__dirname, `./${f}`);
  const success_file = path.join(__dirname, `./result_${f}`);

  let img_arr = [];
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(xlsx_file);
  await workbook.eachSheet((sheet) => {
    sheet.eachRow((row, row_num) => {
      row.eachCell((cell, col_num) => {
        if (cell.value.toString().startsWith("http")) {
          const img_url = cell.value.toString();
          cell.value = null;
          cell.style.alignment.horizontal = "center";
          cell.style.alignment.vertical = "justify";
          const obj = {
            sheet: sheet,
            img_url,
            position: {
              col: col_num,
              row: row_num,
            },
          };
          img_arr.push(obj);
        }
      });
    });
  });

  await saveImg(workbook, img_arr);

  await workbook.xlsx.writeFile(success_file).then(() => {
    console.info("\n 处理成功，程序将在 5s 后自动关闭\n");
    setTimeout(() => {
      process.exit();
    }, 5000);
  });
};

start();
