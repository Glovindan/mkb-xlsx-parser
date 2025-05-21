import * as XLSX from "xlsx";
import * as fs from "fs";
import * as path from "path";

const excelFilePath = path.join("./МКБ-10.xlsx");
const jsonFilePath = path.join("./output.json");

function getJsonData() {
  const workbook = XLSX.readFile(excelFilePath);
  const sheetName = workbook.SheetNames[0];
  const worksheet = workbook.Sheets[sheetName];
  const jsonData: JsonDataItem[] = XLSX.utils.sheet_to_json(worksheet, { defval: null });

  return jsonData;
}

function writeJsonFile(jsonData: any) {
  try {
    fs.writeFileSync(jsonFilePath, JSON.stringify(jsonData, null, 2), "utf-8");
    console.log(`✅ Данные сохранены в файл: ${jsonFilePath}`);
  } catch (error) {
    console.error("Ошибка при записи файла:", error);
    throw new Error()
  }
}

type JsonDataItem = {
  "code": string,
  "description": string,
  "level": number
};
/** Строка таблицы МКБ */
interface MkbTableRow {
  /** Код диагноза */
  code: string,
  /** Название диагноза */
  name: string,
  /** Код родителя */
  parentCode: string,
}

// Парсинг в требуемый формат
function parseJsonData(jsonData: JsonDataItem[]): MkbTableRow[] {
  const jsonDataParsed: MkbTableRow[] = [];

  const codesStack: string[] = [];
  let jsonItemBefore: JsonDataItem | undefined;

  for(const jsonItem of jsonData) {
    const parsedItem: MkbTableRow = {
      code: jsonItem.code,
      name: jsonItem.description,
      parentCode: ""
    }

    if(jsonItemBefore && jsonItem.level > jsonItemBefore.level) {
      // Если уровень увеличился - добавить код предыдущего в стек
      codesStack.push(jsonItemBefore.code)
    } else {
      // Если уровень уменьшился - убрать последней код из стека
      codesStack.pop()
    }

    const parentCode = codesStack.length ? codesStack[codesStack.length - 1] : undefined;
    parsedItem.parentCode = parentCode ?? "";

    jsonItemBefore = jsonItem;
    jsonDataParsed.push(parsedItem);
  }

  return jsonDataParsed;
};

// Получение файла
const jsonData = getJsonData();
// Парсинг в нужный формат
const jsonDataParsed = parseJsonData(jsonData);
// Запись файла
writeJsonFile(jsonDataParsed);