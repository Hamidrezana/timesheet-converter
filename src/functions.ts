import XLSX from "xlsx";
import ExcelJS from "exceljs";
import { TimeSheet, XLSType } from "./types";

export function readXLS(src: string) {
  const workbook = XLSX.readFile(src);
  const firstSheetName = workbook.SheetNames[1];
  const worksheet = workbook.Sheets[firstSheetName];
  const jsonData = XLSX.utils.sheet_to_json(worksheet, {
    raw: false,
  }) as XLSType;

  return jsonData;
}

export function exportTimeSheet(
  jsonData: XLSType,
  startId: string,
  endId: string
) {
  const data: TimeSheet[] = [];
  let shouldRead = false;
  jsonData.forEach((el) => {
    if (el.__EMPTY === startId) {
      shouldRead = true;
    } else if (el.__EMPTY === endId) {
      shouldRead = false;
    } else if (
      shouldRead &&
      el.__EMPTY &&
      (el.__EMPTY as string).includes("/")
    ) {
      data.push({
        dayName: el.__EMPTY_1,
        date: el.__EMPTY,
        start: el.__EMPTY_2,
        end: el.__EMPTY_3,
      });
    }
  });

  return data;
}

function timeFormat(time: string) {
  const [hours, minutes] = time.split(":").map(Number);
  const currentTimeInMinutes = hours * 60 + minutes;

  const newTimeInMinutes = currentTimeInMinutes;
  const newHours = Math.floor(newTimeInMinutes / 60) % 24;
  const newMinutes = newTimeInMinutes % 60;

  return `${String(newHours).padStart(2, "0")}:${String(newMinutes).padStart(
    2,
    "0"
  )}`;
}

export async function importTimeSheet(
  data: XLSType,
  dest: string,
  modifiedDest: string,
  startRow = 5
) {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(dest);

  const worksheet = workbook.getWorksheet(1);
  let row = startRow;

  if (worksheet) {
    data.forEach((el) => {
      if (el.start) {
        const dayCell = worksheet.getCell(`H${row}`);
        if (dayCell.value) {
          const startCell = worksheet.getCell(`I${row}`);
          const endCell = worksheet.getCell(`J${row}`);

          startCell.value = timeFormat(el.start);
          endCell.value = timeFormat(el.end || "00:00");
        }
      }
      row++;
    });

    await workbook.xlsx.writeFile(modifiedDest);
  }
}
