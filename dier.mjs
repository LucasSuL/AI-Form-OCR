import fs from "fs";
import path from "path";

// Define the parent folder directly without using __dirname
const parentFolder = path.join("D:", "02-FullStack", "AI Form OCR", "images");

// Read the contents of the parent folder (images)
try {
  const surveyFolders = fs
    .readdirSync(parentFolder, { withFileTypes: true })
    .filter((dirent) => dirent.isDirectory()) // Filter to get only directories
    .map((dirent) => dirent.name); // Map to get folder names

  // Iterate through each folder and print its name
  for (const folder of surveyFolders) {
    console.log(`Processing ${folder}...`);
  }
} catch (error) {
  console.error("Error reading folders:", error);
}
