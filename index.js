"use strict";

const fs = require("fs");
const xlsx = require("node-xlsx");
const showdown = require("showdown");

function markdownToHtml(markdown) {
  let converter = new showdown.Converter();
  //进行转换
  let html = converter.makeHtml(markdown);
  return html;
}

function htmlToText(html) {
  let text = html
    .replace(/<(p|div)[^>]*>(<br\/?>|&nbsp;)<\/\1>/gi, "\n")
    .replace(/<br\/?>/gi, "\n")
    .replace(/<[^>/]+>/g, "")
    .replace(/(\n)?<\/([^>]+)>/g, "")
    .replace(/\u00a0/g, " ")
    .replace(/&nbsp;/g, " ")
    .replace(/<\/?(img)[^>]*>/gi, "")
    .replace(/&amp;/g, "&")
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&#39;/g, "'")
    .replace(/&quot;/g, '"')
    .replace(/<\/?.+?>/g, "");
  return text;
}

// start end 需要转换开始结束位置
function readExcel(start, end) {
  let current_num = 0;
  let start_time = new Date();
  console.log(`================数据开始处理================`);
  let timer = setInterval(() => {
    console.log(`================数据处理中${current_num}%================`);
    current_num++;
    if (current_num === 100) {
      clearInterval(timer);
    }
  }, 100);

  const excel = xlsx.parse("./origin.xlsx");
  let data = [];
  // excel[0] 第0个工作表
  let excel_content = excel[0].data;
  let title = [];
  let origin_title = excel_content[0];
  // excel_content[0] 第一行 标题
  for (let i in origin_title) {
    title.push(origin_title[i]);
    if (i >= start && i < end) {
      title.push(origin_title[i] + "(html)");
      title.push(origin_title[i] + "(text)");
    }
  }
  data.push(title);
  excel_content.splice(0, 1);
  let origin_content = excel_content;
  // console.log(origin_content);
  origin_content.forEach((item) => {
    let content = [];
    for (let i in item) {
      content.push(item[i]);
      if (i >= start && i < end) {
        content.push(markdownToHtml(item[i]));
        content.push(htmlToText(markdownToHtml(item[i])));
      }
    }
    data.push(content);
  });
  console.log(`================数据处理完毕================`);
  console.log(`================开始导出================`);
  writeXls(data, start_time, timer);
}

function writeXls(datas, start_time, timer) {
  let buffer = xlsx.build([
    {
      name: "sheet0",
      data: datas,
    },
  ]);
  fs.access("./result.xlsx", (err) => {
    if (err) {
      fs.writeFileSync("./result.xlsx", buffer, { flag: "w" });
      let end_time = new Date();
      let duration = end_time - start_time;
      console.log(`================数据处理中${100}%================`);
      console.log(
        `================导出完毕用时：${duration}ms================`
      );
      clearInterval(timer);
    } else {
      // fs.unlink删除文件
      fs.unlink("./result.xlsx", function (error) {
        if (error) {
          console.log(error);
        }
        fs.writeFileSync("./result.xlsx", buffer, { flag: "w" });
        let end_time = new Date();
        let duration = end_time - start_time;
        console.log(`================数据处理中${100}%================`);
        console.log(
          `================导出完毕用时：${duration}ms================`
        );
        clearInterval(timer);
      });
    }
  });
}

readExcel(6, 12);
