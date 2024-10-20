import OpenAI from "openai";
import fs from "fs";
import ExcelJS from "exceljs";

const openai = new OpenAI({
  apiKey:process.env.GPT_API
});

// Function to convert image to Base64
function convertImageToBase64(imagePath) {
  const imageBuffer = fs.readFileSync(imagePath);
  const base64Image = imageBuffer.toString("base64");
  return base64Image;
}

function cleanString(inputString) {
  // 将字符串按行拆分为数组
  const lines = inputString.trim().split("\n");
  return lines.slice(1, -1).join('\n'); // 去掉第一行和最后一行，并重新拼接为字符串
}

// // 展平 JSON 对象
// const flattenObject = (obj, prefix = '') => {
//     return Object.keys(obj).reduce((acc, key) => {
//         const pre = prefix.length ? `${prefix}.` : '';
//         if (typeof obj[key] === 'object' && !Array.isArray(obj[key])) {
//             Object.assign(acc, flattenObject(obj[key], pre + key));
//         } else {
//             acc[pre + key] = obj[key];
//         }
//         return acc;
//     }, {});
// };


async function main() {
  const base64Image1 = convertImageToBase64("input1.jpg"); 
  const base64Image2 = convertImageToBase64("input2.jpg"); 
  const base64Image3 = convertImageToBase64("input3.jpg"); 
  const base64Image4 = convertImageToBase64("input4.jpg"); 
  const base64Image5 = convertImageToBase64("input5.jpg"); 
  const base64Image6 = convertImageToBase64("input6.jpg"); 

  const prompt = `
    From the hand-written survey images, identify which selection the user has chosen. The expected output should be in the following JSON format, this is in order with images:
    {
      "1 bedroom": 1 | 0,
          "1 bedroom + study": 1 | 0,
          "2 bedrooms": 1 | 0,
          "2 bedrooms + study": 1 | 0,
          "3 bedrooms": 1 | 0,
          "3 bedrooms + study": 1 | 0,
          "1 bathroom": 1 | 0,
          "1 bathroom, 2 toilets": 1 | 0,
          "2 bathrooms": 1 | 0,
          "2 bathrooms, 2 toilets": 1 | 0,
          "Courtyard": 1 | 0,
          "Open balcony": 1 | 0,
          "Enclosed balcony (Wintergarden)": 1 | 0,
          "Either": 1 | 0,

          "1 car": 1 | 2 | 3 | 4 | 5,
        "2 cars": 1 | 2 | 3 | 4 | 5,
        "Electric car charging": 1 | 2 | 3 | 4 | 5,
        "Electric bicycle": 1 | 2 | 3 | 4 | 5,
        "Mobility scooter": 1 | 2 | 3 | 4 | 5,

         "Natural ventilation for every residence": 1 | 2 | 3 | 4 | 5,
        "Solar generated power": 1 | 2 | 3 | 4 | 5,
        "Rainwater harvesting system": 1 | 2 | 3 | 4 | 5,
        "Grey-water recycling system": 1 | 2 | 3 | 4 | 5,
        "Designed for energy efficiency": 1 | 2 | 3 | 4 | 5,
        "Energy efficient appliances": 1 | 2 | 3 | 4 | 5,
        "Double glazing windows": 1 | 2 | 3 | 4 | 5,
        "Reverse-cycle air-conditioned zones": 1 | 2 | 3 | 4 | 5,

        "Car wash bay": 1 | 2 | 3 | 4 | 5,
        "Secure storage": 1 | 2 | 3 | 4 | 5,
        "Wine cellar": 1 | 2 | 3 | 4 | 5,
        "Hairdressing Salon": 1 | 2 | 3 | 4 | 5,
        "Beauty Salon": 1 | 2 | 3 | 4 | 5,
        "Meeting rooms": 1 | 2 | 3 | 4 | 5,
        "Coffee machine (residents)": 1 | 2 | 3 | 4 | 5,
        "Business Centre": 1 | 2 | 3 | 4 | 5,
        "Darts": 1 | 2 | 3 | 4 | 5,
        "Billiard table": 1 | 2 | 3 | 4 | 5,
        "Private dining room (family occasions/entertaining)": 1 | 2 | 3 | 4 | 5,
        "Wi-Fi": 1 | 2 | 3 | 4 | 5,
        "Pizza Oven": 1 | 2 | 3 | 4 | 5,
        "Bocce/Pétanque": 1 | 2 | 3 | 4 | 5,
        "Other amenities/services": String | None,

        "Laundry services": 1 | 2 | 3 | 4 | 5,
        "Shopping": 1 | 2 | 3 | 4 | 5,
        "Nursing Services": 1 | 2 | 3 | 4 | 5,
        "General and regular household cleaning": 1 | 2 | 3 | 4 | 5,
        "Spring cleaning": 1 | 2 | 3 | 4 | 5,
        "Window cleaning (external)": 1 | 2 | 3 | 4 | 5,
        "Wellness checks": 1 | 2 | 3 | 4 | 5,
        "Personal Trainer/Exercise Physiologist": 1 | 2 | 3 | 4 | 5,
        "Receiving online shopping (groceries)": 1 | 2 | 3 | 4 | 5,
        "Receiving online shopping (non-perishable)": 1 | 2 | 3 | 4 | 5,
        "Housekeeping services": 1 | 2 | 3 | 4 | 5,
        "Medication management": 1 | 2 | 3 | 4 | 5,
        "Lifestyle/social coordinator": 1 | 2 | 3 | 4 | 5,
        "Meals delivered": 1 | 2 | 3 | 4 | 5,
        "Personal care": 1 | 2 | 3 | 4 | 5,
        "Bus transport service": 1 | 2 | 3 | 4 | 5,
        "Concierge": 1 | 2 | 3 | 4 | 5,
        "Waste/garbage collection": 1 | 2 | 3 | 4 | 5,
        "Other extra service (please list)": String | None

        Rules: just return the string from '{' to '}'.
    `;

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
          // {
          //   type: "image_url",
          //   image_url: {
          //     url: `data:image/jpeg;base64,${base64Image1}`,
          //   },
          // },
          {
            type: "image_url",
            image_url: {
              url: `data:image/jpeg;base64,${base64Image2}`,
            },
          },
          // {
          //   type: "image_url",
          //   image_url: {
          //     url: `data:image/jpeg;base64,${base64Image3}`,
          //   },
          // },
          // {
          //   type: "image_url",
          //   image_url: {
          //     url: `data:image/jpeg;base64,${base64Image4}`,
          //   },
          // },
          {
            type: "image_url",
            image_url: {
              url: `data:image/jpeg;base64,${base64Image5}`,
            },
          },
          // {
          //   type: "image_url",
          //   image_url: {
          //     url: `data:image/jpeg;base64,${base64Image6}`,
          //   },
          // },
        ],
      },
    ],
  });

  const jsonData = response.choices[0].message.content;
  console.log("jsonData:");
  console.log(jsonData);

  const cleanedString = cleanString(jsonData);
  console.log("cleanedString:");

  console.log(cleanedString);
  let jsonObject = null;

  try {
    // Convert the JSON string to a JavaScript object
    jsonObject = JSON.parse(cleanedString);

    // Log the result
    // console.log(jsonObject);
  } catch (error) {
    console.error("Error parsing JSON:", error);
  }

  // 创建一个新的工作簿
  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet("My Sheet");

  // 设置表头
  worksheet.columns = Object.keys(jsonObject).map((key) => ({
    header: key,
    key,
  }));

  // 将 JSON 对象数据写入工作表
  worksheet.addRow(jsonObject);

  // 保存 Excel 文件
  workbook.xlsx
    .writeFile("output.xlsx")
    .then(() => {
      console.log("Excel file created successfully!");
    })
    .catch((err) => {
      console.error("Error creating Excel file:", err);
    });
}
main();
