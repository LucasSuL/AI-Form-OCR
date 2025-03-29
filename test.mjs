import OpenAI from "openai";
import fs from "fs";
import ExcelJS from "exceljs";
import path from "path";

const openai = new OpenAI({
  apiKey: process.env.GPT_API,
});

const parentFolder = path.join("D:", "11-Survey PDFs", "Siquilla", "PDFs");

// 从 prompt.txt 读取内容
const prompt = fs.readFileSync(path.join(".", "prompt-0323.txt"), "utf-8");

function convertImageToBase64(imagePath) {
  const imageBuffer = fs.readFileSync(imagePath);
  return imageBuffer.toString("base64");
}

function cleanString(inputString) {
  // 去掉 Markdown 的 ```json 或 ``` 包裹行
  const lines = inputString.trim().split("\n");

  // 如果首行是 ```json 或 ```，就移除首尾行
  let body = lines;
  if (lines[0].startsWith("```")) {
    body = lines.slice(1, -1); // 移除代码块的三引号
  }

  // 去除所有反引号（有些字段内容也会出现 `xxx`）
  const joined = body.join("\n").replace(/`/g, "");

  return joined;
}

// 统一格式处理函数
function normalizeField(key, value) {
  if (typeof value !== "string") return value;

  // 姓名字段：首字母大写
  if (["Last Name/s", "First Name/s"].includes(key)) {
    return value
      .toLowerCase()
      .split(/\s+/)
      .map((w) => w.charAt(0).toUpperCase() + w.slice(1))
      .join(" ");
  }

  // 地址字段：Title Case
  if (key === "Address") {
    return value
      .toLowerCase()
      .split(" ")
      .map((w) => w.charAt(0).toUpperCase() + w.slice(1))
      .join(" ");
  }

  // 邮箱统一小写
  if (key === "Email") {
    return value.toLowerCase();
  }

  // 手机号码：加前导 0（如果是 9 位数字）
  if (key === "Mobile") {
    const digits = value.replace(/\D/g, "");
    return digits.length === 9 ? "0" + digits : digits;
  }

  // 自由文本：首字母大写，其余小写
  if (value.length > 0 && /^[a-z]/i.test(value.charAt(0))) {
    return value.charAt(0).toUpperCase() + value.slice(1).toLowerCase();
  }

  return value;
}

async function main() {
  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet("Survey Results");
  let headerWritten = false;

  try {
    const surveyFolders = fs
      .readdirSync(parentFolder, { withFileTypes: true })
      .filter((dirent) => dirent.isDirectory())
      .map((dirent) => dirent.name)
      .sort((a, b) => {
        const numA = Number(a.replace(/[^0-9]/g, ""));
        const numB = Number(b.replace(/[^0-9]/g, ""));
        return numA - numB;
      });

    for (const folder of surveyFolders) {
      console.log(`Processing ${folder}...`);
      const folderPath = path.join(parentFolder, folder);

      const imageFiles = fs
        .readdirSync(folderPath)
        .filter((file) => file.endsWith(".jpg") || file.endsWith(".jpeg"))
        .sort();

      const imageFilesToUse = imageFiles.slice(1); // 跳过封面

      if (imageFilesToUse.length === 0) {
        console.warn(`No images to process in ${folder}`);
        continue;
      }

      const imagePaths = imageFilesToUse.map((image) =>
        path.join(folderPath, image)
      );

      const imageDataArray = imagePaths.map(convertImageToBase64);

      const response = await openai.chat.completions.create({
        model: "gpt-4o",
        messages: [
          {
            role: "user",
            content: [
              {
                type: "text",
                text: prompt,
              },
              ...imageDataArray.map((base64Image) => ({
                type: "image_url",
                image_url: { url: `data:image/jpeg;base64,${base64Image}` },
              })),
            ],
          },
        ],
      });

      const jsonData = response.choices[0].message.content;
      const cleanedString = cleanString(jsonData);

      let jsonObject = null;
      try {
        jsonObject = JSON.parse(cleanedString);
      } catch (error) {
        console.error("Error parsing JSON:", error);
        continue;
      }

      if (!headerWritten) {
        worksheet.columns = [
          { header: "Folder Name", key: "folderName" },
          ...Object.keys(jsonObject).map((key) => ({
            header: key,
            key,
          })),
        ];
        headerWritten = true;
      }

      const normalizedRow = {};
      for (const [key, value] of Object.entries(jsonObject)) {
        normalizedRow[key] = normalizeField(key, value);
      }

      worksheet.addRow({
        folderName: folder.replace(/\.pdf$/i, ""),
        ...normalizedRow,
      });

      console.log(`OCR success for ${folder}!`);
    }

    await workbook.xlsx.writeFile(
      path.join(parentFolder, "survey_data_sup.xlsx")
    );
    console.log("Excel file saved!");
  } catch (error) {
    console.error("Error during processing:", error);
  }
}

main();
