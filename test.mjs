import OpenAI from "openai";
import fs from "fs";
import ExcelJS from "exceljs";
import path from "path";

const openai = new OpenAI({
  apiKey: process.env.GPT_API,
});

const parentFolder = path.join("D:", "11-Survey PDFs", "root folder");

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
  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet("Survey Results");
  let headerWritten = false;

    const prompt = `
   As an expert-level administrative assistant in the administrative department, your task is to create a data entry document that improves efficiency and accuracy in data management. The output should be a high-quality document with organized and error-free data entry. The finished work will be used by the team for data analysis and decision-making. Core success factors include attention to detail, timeliness, and accuracy, measured by the document's ability to streamline data entry processes and reduce errors.

     Rules:
      1- If an option is ticked, represent it as '1'; otherwise, represent it as '0'.
      2- When a number is circled, return that number, which should be between 1 and 5. The page will be divided into two columns, with 1 on the leftmost side and 5 on the rightmost side. If you cannot recognise the number, use its position to determine the corresponding value. If the number is uncertain, return 3
      3- If a string to be returned is not written, or no number is circled, just leave it empty.
      4- Do NOT contain backtick within brackets since it will cause parsing error.

    The expected JSON output format is as follows:
    {
         "Last Name/s": String,  // e.g., "Smith"
        "First Name/s": String,  // e.g., "John"
        "Address": String,       // e.g., "123 Example St, Adelaide"
        "Email": String,         // e.g., "john.smith@example.com"
        "Mobile": String,        // e.g., "0412345678" without space between
          "Interested in Westering for yourself as a resident": 1 | 0,
              "Family": 1 | 0,
              "Friend": 1 | 0,
              "Other person": String,
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
            "Other person live with me": String,
          "Local Community Membership": 1 | 0,
            "If you have Local Community Membership, please list": String,
          "Pets": 1 | 0,
          "If you have pets, what type:": String,
            "Being part of a community": 1 | 0,
            "Security": 1 | 0,
            "Staying independent and connected to friends": 1 | 0,
            "Services available": 1 | 0,
            "Low/no maintenance worries": 1 | 0,
            "Other reason of retirement living": String,
            "Active and engaged with friends": 1 | 0,
            "Socially independent": 1 | 0,
            "Desire for more engagement": 1 | 0,
            "Other lifestyle": String,
            "Totally independent": 1 | 0,
            "Private services": 1 | 0,
            "Home Care package services": 1 | 0,
            "Family support": 1 | 0,
            "Apartment": 1 | 0,
            "Villa/House": 1 | 0,
            "Either housing style": 1 | 0,
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
            "Either courtyard or balcony": 1 | 0,

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

         "Floorboards: Bedrooms": 1 | 0,
          "Floorboards: kitchen": 1 | 0,
          "Floorboards: Living/dining": 1 | 0,
          "Carpet: Bedrooms": 1 | 0,
          "Carpet: Living/dining": 1 | 0,
          "Tiles: Bathroom": 1 | 0,
          "Tiles: Bedrooms": 1 | 0,
          "Tiles: Kitchen": 1 | 0,
          "Tiles: Living/dining": 1 | 0,
          "Underfloor heating: Bathroom": 1 | 0,
          "Underfloor heating: Bedrooms": 1 | 0,
          "Underfloor heating: Living/dining": 1 | 0,

          "Designed for ageing in place": 1 | 2 | 3 | 4 | 5,
          "Level access/no threshold": 1 | 2 | 3 | 4 | 5,
          "Light filled home": 1 | 2 | 3 | 4 | 5,
          "Reverse cycle air conditioning": 1 | 2 | 3 | 4 | 5,
          "Ceiling fan": 1 | 2 | 3 | 4 | 5,
          "Sensor lights": 1 | 2 | 3 | 4 | 5,
          "Electric automated blinds": 1 | 2 | 3 | 4 | 5,
          "European kitchen appliances": 1 | 2 | 3 | 4 | 5,
          "Built-in robes": 1 | 2 | 3 | 4 | 5,
          "Walk-in robe": 1 | 2 | 3 | 4 | 5,
          "Sliding robe doors": 1 | 2 | 3 | 4 | 5,
          "Open-out robe doors": 1 | 2 | 3 | 4 | 5,
          "No robe doors": 1 | 2 | 3 | 4 | 5,
          "Linen cupboard": 1 | 2 | 3 | 4 | 5,
          "Open-plan living area": 1 | 2 | 3 | 4 | 5,
          "Built-in coffee machine": 1 | 2 | 3 | 4 | 5,
          "Built-in microwave oven": 1 | 2 | 3 | 4 | 5,
                 "Pantry storage": 1 | 2 | 3 | 4 | 5,
                 "Butler's pantry": 1 | 2 | 3 | 4 | 5,
          "Steam oven": 1 | 2 | 3 | 4 | 5,
          "Warming drawer": 1 | 2 | 3 | 4 | 5,
          "Dishwasher - full size": 1 | 2 | 3 | 4 | 5,
          "Dishwasher - drawer (half size)": 1 | 2 | 3 | 4 | 5,
          "Filtered water": 1 | 2 | 3 | 4 | 5,
          "Zip Tap": 1 | 2 | 3 | 4 | 5,

          "Plumbed fridge": 1 | 2 | 3 | 4 | 5,
          "Wine fridge": 1 | 2 | 3 | 4 | 5,
          "Shower with flexible hose": 1 | 2 | 3 | 4 | 5,
          "European laundry": 1 | 2 | 3 | 4 | 5,
          "Laundry in bathroom": 1 | 2 | 3 | 4 | 5,
          "Separate laundry": 1 | 2 | 3 | 4 | 5,
          "Combined washer/dryer": 1 | 2 | 3 | 4 | 5,
          "Raised laundry appliances": 1 | 2 | 3 | 4 | 5,
          "Laundry trough": 1 | 2 | 3 | 4 | 5,
          "Ironing board storage": 1 | 2 | 3 | 4 | 5,
          "Stick vacuum storage": 1 | 2 | 3 | 4 | 5,
          "Bath tub": 1 | 2 | 3 | 4 | 5,
          "Electric fireplace in living area": 1 | 2 | 3 | 4 | 5,
          "Smart home features": 1 | 2 | 3 | 4 | 5,
          "Small safe in discreet area": 1 | 2 | 3 | 4 | 5,
          "Keyless entry (e.g., fob)": 1 | 2 | 3 | 4 | 5,
          "Intercom for visitors - audio only": 1 | 2 | 3 | 4 | 5,
          "Intercom for visitors - visual & audio": 1 | 2 | 3 | 4 | 5,
          "Heated towel rails": 1 | 2 | 3 | 4 | 5,
          "Other appliance, list here:": String,

          "24/7 Emergency Call System": 1 | 2 | 3 | 4 | 5,
          "Emergency call buttons installed": 1 | 2 | 3 | 4 | 5,
          "Emergency call pendant": 1 | 2 | 3 | 4 | 5,
          "Secure garage access": 1 | 2 | 3 | 4 | 5,
          "Walking distance to shops": 1 | 2 | 3 | 4 | 5,
          "Walking distance to cafe": 1 | 2 | 3 | 4 | 5,
          "Walking distance to park": 1 | 2 | 3 | 4 | 5,
          "Views/aspects": 1 | 2 | 3 | 4 | 5,
          "Easy access to public transport (bus)": 1 | 2 | 3 | 4 | 5,
          "Being part of a like-minded community": 1 | 2 | 3 | 4 | 5,
          "Co-located to aged care": 1 | 2 | 3 | 4 | 5,
          "Low-maintenance living": 1 | 2 | 3 | 4 | 5,
          "Gardens maintained by Helping Hand": 1 | 2 | 3 | 4 | 5,
          "Building maintained by Helping Hand": 1 | 2 | 3 | 4 | 5,
          "No AirBnB/vacation rental neighbours": 1 | 2 | 3 | 4 | 5,
          "Playground equipment": 1 | 2 | 3 | 4 | 5,
          "Secure access to buildings": 1 | 2 | 3 | 4 | 5,
          "Security patrols": 1 | 2 | 3 | 4 | 5,

          "Library": 1 | 2 | 3 | 4 | 5,
          "Herb or vegetable garden": 1 | 2 | 3 | 4 | 5,
          "Barbeque areas": 1 | 2 | 3 | 4 | 5,
          "Table Tennis": 1 | 2 | 3 | 4 | 5,
          "Wellness/Exercise classes": 1 | 2 | 3 | 4 | 5,
          "Club lounge area": 1 | 2 | 3 | 4 | 5,
          "Fitness Centre ": 1 | 2 | 3 | 4 | 5,
          "Communal Hobby Workshop": 1 | 2 | 3 | 4 | 5,
          "Green Space/Gardens": 1 | 2 | 3 | 4 | 5,
          "Heated swimming pool": 1 | 2 | 3 | 4 | 5,
          "Sauna": 1 | 2 | 3 | 4 | 5,
          "Village/community bus": 1 | 2 | 3 | 4 | 5,
          "Art and Craft Studio": 1 | 2 | 3 | 4 | 5,
          "Cafe": 1 | 2 | 3 | 4 | 5,
          "Shop/providore": 1 | 2 | 3 | 4 | 5,
          "Bar": 1 | 2 | 3 | 4 | 5,
          "Cinema": 1 | 2 | 3 | 4 | 5,
          "Restaurant": 1 | 2 | 3 | 4 | 5,
          "Visiting allied health practitioners": 1 | 2 | 3 | 4 | 5,
          "Pharmacy": 1 | 2 | 3 | 4 | 5,
          "Spa": 1 | 2 | 3 | 4 | 5,

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
          "Private dining room": 1 | 2 | 3 | 4 | 5,
          "Wi-Fi": 1 | 2 | 3 | 4 | 5,
          "Pizza Oven": 1 | 2 | 3 | 4 | 5,
          "Bocce/Pétanque": 1 | 2 | 3 | 4 | 5,
          "Other amenities/services": String,

          "Linen services": 1 | 2 | 3 | 4 | 5,
          "Shopping": 1 | 2 | 3 | 4 | 5,
          "Nursing Services": 1 | 2 | 3 | 4 | 5,
          "General and regular household cleaning": 1 | 2 | 3 | 4 | 5,
          "Spring cleaning": 1 | 2 | 3 | 4 | 5,
          "Window cleaning (external)": 1 | 2 | 3 | 4 | 5,
          "Window cleaning (internal)": 1 | 2 | 3 | 4 | 5,
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
          "Other extra service (please list)": String,

          "Less than 1 year": 1 | 0,
          "1 to 2 years": 1 | 0,
          "2 to 4 years": 1 | 0,
          "4+ years": 1 | 0,

          "Is there anything else you would like to share with us that you believe will be important to consider in our planning for your next home and community": String,
          "And finally, please tell us what is most appealing for you about our plans for Westering North Adelaide?": String,
          }
          Please try your best.
      `;

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
        .filter((file) => file.endsWith(".jpg") || file.endsWith(".jpeg")); // Filter for JPG/JPEG files

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
                image_url: { url: `data:image/jpeg;base64,${base64Image}` },
              })),
            ],
          },
        ],
      });

      const jsonData = response.choices[0].message.content;
      console.log(`OCR success for ${folder}!`);
      // console.log("jsonData:");
      // console.log(jsonData);

      const cleanedString = cleanString(jsonData);
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
