import OpenAI from "openai";
import fs from "fs";
import ExcelJS from "exceljs";

const openai = new OpenAI({
  apiKey:process.env.GPT_API
});

// Function to convert image to Base64
function convertImageToBase64(imagePath) {
  // Read the image file synchronously
  const imageBuffer = fs.readFileSync(imagePath);
  // Convert the buffer to a Base64 string
  const base64Image = imageBuffer.toString("base64");
  return base64Image;
}

function cleanString(inputString) {
  // 按行分割字符串
  const lines = inputString.split("\n");

  // 去掉第一行和最后一行
  const cleanedLines = lines.slice(1, -1);

  // 重新拼接成字符串
  return cleanedLines.join("\n");
}

// 展平 JSON 对象
const flattenObject = (obj, prefix = '') => {
    return Object.keys(obj).reduce((acc, key) => {
        const pre = prefix.length ? `${prefix}.` : '';
        if (typeof obj[key] === 'object' && !Array.isArray(obj[key])) {
            Object.assign(acc, flattenObject(obj[key], pre + key));
        } else {
            acc[pre + key] = obj[key];
        }
        return acc;
    }, {});
};


async function main() {
  const base64Image1 = convertImageToBase64("input1.jpg"); // Replace with the path to your local image
  const base64Image2 = convertImageToBase64("input2.jpg"); // Replace with the path to your local image
  // const base64Image3 = convertImageToBase64("input3.jpg"); // Replace with the path to your local image
  const base64Image7 = convertImageToBase64("input7.jpg"); // Replace with the path to your local image

  const prompt = `
    From the hand-written survey images, identify which selection the user has chosen. The expected output should be in the following JSON format:
    {
      "Last Name/s": String,
      "First Name/s": String,
      "Address": String,
      "Email": String,
      "Mobile": Number,
        "Interested in Westering North Adelaide as a resident": 1 | 0,
            "Family": 1 | 0,
            "Friend": 1 | 0,
            "Other person": String | None,
          "Male": 1 | 0,
          "Female": 1 | 0,
          "Single": 1 | 0,
          "Couple": 1 | 0,
          "55 to 59 years": 1 | 0,
          "60 to 64 years": 1 | 0,
          "65 to 69 years": 1 | 0,
          "70 to 74 years": 1 | 0,
          "75 to 79 years": 1 | 0,
          "80 to 84 years": 1 | 0,
          "85 to 89 years": 1 | 0,
          "90+ years": 1 | 0,
          "I live alone": 1 | 0,
          "My partner": 1 | 0,
          "My son/daughter": 1 | 0,
          "Other": 1 | 0,
        "Local Community Membership": 1 | 0,
          "If you have Local Community Membership, please list": String | None,
        "Pets": 1 | 0,
        "If you have pets, what type:": String | None,
          "Being part of a community": 1 | 0,
          "Security": 1 | 0,
          "Staying independent and connected to friends": 1 | 0,
          "Services available": 1 | 0,
          "Low/no maintenance worries": 1 | 0,
          "Other reason of retirement living": String | None,
          "Active and engaged with friends": 1 | 0,
          "Socially independent": 1 | 0,
          "Desire for more engagement": 1 | 0,
          "Other lifestyle": String | None,
          "Totally independent": 1 | 0,
          "Private services (e.g. cleaning, gardening)": 1 | 0,
          "Home Care package services": 1 | 0,
          "Family support": 1 | 0,
          "Apartment": 1 | 0,
          "Villa/House": 1 | 0,
          "Either": 1 | 0,
          "Less than $500,000": 1 | 0,
          "$500,000 - $1m": 1 | 0,
          "$1m - $1.5m": 1 | 0,
          "$1.5m - $2m": 1 | 0,
          "$2m - $2.5m": 1 | 0,
          "Above $2.5m": 1 | 0,
          "Need to Sell Home": 1 | 0,

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

        "Less than 1 year": 1 | 0,
        "1 to 2 years": 1 | 0,
        "2 to 4 years": 1 | 0,
        "4+ years": 1 | 0,

        "Is there anything else you would like to share with us that you believe will be important to consider in our planning for your next home and community": String | None,

        "And finally, please tell us what is most appealing for you about our plans for Westering North Adelaide?": String | None
        


        }
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
          {
            type: "image_url",
            image_url: {
              url: `data:image/jpeg;base64,${base64Image1}`, // Use the Base64 encoded image here
            },
          },
          {
            type: "image_url",
            image_url: {
              url: `data:image/jpeg;base64,${base64Image2}`, // Use the Base64 encoded image here
            },
          },
          // {
          //   type: "image_url",
          //   image_url: {
          //     url: `data:image/jpeg;base64,${base64Image3}`, // Use the Base64 encoded image here
          //   },
          // },
          {
            type: "image_url",
            image_url: {
              url: `data:image/jpeg;base64,${base64Image7}`, // Use the Base64 encoded image here
            },
          },
        ],
      },
    ],
  });


  const jsonData = response.choices[0].message.content;
  const cleanedString = cleanString(jsonData);

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
