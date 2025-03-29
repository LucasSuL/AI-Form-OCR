import OpenAI from "openai";
import fs from "fs";
import ExcelJS from "exceljs";
import path from "path";

const openai = new OpenAI({
  apiKey: process.env.GPT_API,
});

const parentFolder = path.join("D:", "11-Survey PDFs", "da folder");

function convertImageToBase64(imagePath) {
  const imageBuffer = fs.readFileSync(imagePath);
  return imageBuffer.toString("base64");
}

function cleanString(inputString) {
  const lines = inputString.trim().split("\n");
  return lines.slice(1, -1).join("\n"); // remove first and last line, and conbine again
}

async function main() {
  // create new excel
  // const workbook = new ExcelJS.Workbook();
  // const worksheet = workbook.addWorksheet("Survey Results");
  // let headerWritten = false;
  const prompt = `You are a expert consultant for a retirement living village operation team, and your company is going to build a new place.
  Read the images of potential client survey, and give me a summary to show the executive team how they should design their new community to meet potential clients' expectations
`;

//     const prompt = `You are a expert consultant for a retirement living village operation team, and your company is going to build a new place.
//     Read the images of data analysis regarding the surveys, and reponse a report containing both high-level summary for each category and detailed insights for each category, to show the executive team how they should design their new community using this reference or similar ones:
    
//     reference: 
//     Age-Based Distribution: The bar chart shows a peak in the age range of [X-Y years], suggesting that this demographic is the primary user base. There is a sharp drop in engagement or presence beyond [Age Z].
// Insight: This skew in age distribution indicates that the primary audience is younger (or older) and likely driven by factors such as accessibility, price sensitivity, or preferences. Further analysis could explore how these factors impact user behaviour or engagement.`;

  try {
    const surveyFolders = fs
      .readdirSync(parentFolder, { withFileTypes: true })
      .filter((dirent) => dirent.isDirectory()) // Filter to get only directories
      .map((dirent) => dirent.name); // Map to get folder names

    const imagesByFolder = {};

    // Iterate through each folder and print its name
    for (const folder of surveyFolders) {
      console.log(`Processing ${folder}...`);
      const folderPath = path.join(parentFolder, folder); // Path to the current folder

      // Read all files in the current folder
      const imageFiles = fs
        .readdirSync(folderPath)
        .filter((file) => file.endsWith(".jpg") || file.endsWith(".jpeg")|| file.endsWith(".png")); // Filter for JPG/JPEG files

      // Store the image files in the object
      imagesByFolder[folder] = imageFiles.map((image) =>
        path.join(folderPath, image)
      ); // Full path to images

      // 打印仅文件名
      const imageNames = imageFiles.map((file) => path.basename(file)); // 提取文件名
      console.log(`Found images: ${imageNames.join(", ")}`);

      // 将图像转换为 Base64
      const imageDataArray = imagesByFolder[folder].map(convertImageToBase64);

      // 发送请求到 OpenAI
      const response = await openai.chat.completions.create({
        model: "gpt-4o",
        messages: [
          {
            role: "user",
            content: [
              {
                type: "text",
                text: prompt, // 替换为你的提示
              },
              ...imageDataArray.map((base64Image) => ({
                type: "image_url",
                image_url: { url: `data:image/png;base64,${base64Image}` },
              })),
            ],
          },
        ],
      });

      const jsonData = response.choices[0].message.content;
      console.log(`OCR success for ${folder}!`);
      console.log("jsonData:");
      console.log(jsonData);

      // const cleanedString = cleanString(jsonData);
      let jsonObject = null;

      // console.log("cleanedString:");
      // console.log(cleanedString);

      try {
        jsonObject = JSON.parse(cleanedString);
      } catch (error) {
        console.error("Error parsing JSON:", error);
        continue;
      }

      if (!headerWritten) {
        // set excel header
        worksheet.columns = Object.keys(jsonObject).map((key) => ({
          header: key,
          key,
        }));

        headerWritten = true;
      }

      // 将解析后的 JSON 对象添加到工作表
      worksheet.addRow(jsonObject);
    }
    // 保存 Excel 文件
    await workbook.xlsx.writeFile(path.join(parentFolder, "survey_data1.xlsx"));
    console.log("Excel file saved!");

  } catch (error) {
    console.error("Error reading folders:", error);
  }
}

main();
