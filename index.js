const fs = require("fs/promises");
const path = require("path");
const XLSX = require("xlsx");

const books = "./books";
const jsons = "./jsons";
const recaudados = "./recaudados";
const recaudadosJSON = "./recaudados_json";
const jsonsFinal = './jsonsFinal';
const finalPath = './booksfinal';
const books2json = async (folderPath, output) => {
  try {
    const files = await fs.readdir(folderPath);

    for (const file of files) {
      const filePath = path.join(folderPath, file);

      if (path.extname(file).toLowerCase() === ".xlsx") {
        const workbook = XLSX.readFile(filePath);
        const sheetName = workbook.SheetNames[0];
        const jsonData = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName]);

        const jsonFileName = path.basename(file, ".xlsx") + ".json";
        const jsonFilePath = path.join(output, jsonFileName);
        await fs.writeFile(jsonFilePath, JSON.stringify(jsonData));

        console.log(`Archivo convertido a JSON: ${jsonFilePath}`);
      }
    }
  } catch (err) {
    console.error("Error al leer el directorio:", err);
  }
};

const deleteRepeated = async () => {
  let fileName;
  try {
    const recaudados = await fs.readdir(recaudadosJSON);
    const files = await fs.readdir(jsons);
    for (let recaudo of recaudados) {
      fileName = recaudo;
      const fileContent = await fs.readFile(
        path.join(recaudadosJSON, recaudo),
        "utf-8"
      );
      recaudo = JSON.parse(fileContent);
      for (let file of files) {
        const dataUtil = [];
        const fileRead = await fs.readFile(path.join(jsons, file), {
          encoding: "utf-8",
        });
        const fileParsed = JSON.parse(fileRead);
        for (let client of fileParsed) {
          const foundCoincidence = recaudo.find(
            (d) => d["Cédula/RUC"] === client["Cedula"]
          );
          if (!foundCoincidence) {
            const strClient = {};
            for (let key in client) {
              strClient[key] = String(client[key]).trim();
            }
            dataUtil.push(strClient);
          }
        }
        const finalJSON = JSON.stringify(dataUtil);
        const finalPath = path.join(
          "./jsonsfinal",
          "FINAL-" + path.basename(file)
        );
        await fs.writeFile(finalPath, finalJSON, "utf-8");
        console.log(`Archivo final de ${finalPath} generado correctamente`);
      }
    }
  } catch (err) {
    console.error("Error al leer el directorio:", fileName, err);
  }
};

const convertJSONtoXLSX = async (jsonPath, outputPath) => {
  try {
    const jsonData = await fs.readFile(jsonPath, "utf-8");
    const data = JSON.parse(jsonData);
    const workbook = XLSX.utils.book_new();
    const sheet = XLSX.utils.json_to_sheet(data);
    XLSX.utils.book_append_sheet(workbook, sheet, "Sheet1");
    const outputFilePath = path.join(outputPath, `${path.basename(jsonPath, ".json")}.xlsx`);
    await fs.writeFile(outputFilePath, XLSX.write(workbook, { type: "buffer", bookType: "xlsx" }));
    console.log(`Archivo final XLSX generado correctamente: ${outputFilePath}`);
  } catch (error) {
    console.error("Error al convertir el archivo JSON:", error);
  }
};

const jsonToXLSX = async (folderPath, outputPath) => {
  try {
    const files = await fs.readdir(folderPath);
    for (const file of files) {
      const filePath = path.join(folderPath, file);
      if (path.extname(file).toLowerCase() === ".json") {
        await convertJSONtoXLSX(filePath, outputPath);
      }
    }
  } catch (error) {
    console.error("Error al leer la carpeta de archivos JSON:", error);
  }
};


async function main() {
  await books2json(books, jsons);
  await books2json(recaudados, recaudadosJSON);
  await deleteRepeated();
  await jsonToXLSX(jsonsFinal, finalPath);
}

main();
