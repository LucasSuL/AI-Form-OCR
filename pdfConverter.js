// import { fromPath } from "pdf2pic";
// import { mkdirsSync, writeFileSync } from "fs-extra";
// import { rimraf } from "rimraf";

// export default () => {
//   const specimen1 = "./original/Q1.pdf";
//   const outputDirectory = "./images";
//   rimraf.sync(outputDirectory);
//   mkdirsSync(outputDirectory);

//   const baseOptions = {
//     width: 594,
//     height: 841,
//     density: 330,
//     savePath: outputDirectory,
//   };

//   const convert = fromPath(specimen1, baseOptions);

//   return convert(1, true).then((output) => {
//     writeFileSync(outputDirectory + "/base64-output.txt", output.base64);

//     writeFileSync(
//       outputDirectory + "/base64-output.png",
//       output.base64,
//       "base64"
//     );
//   });
// };
