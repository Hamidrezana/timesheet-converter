import { exportTimeSheet, importTimeSheet, readXLS } from "./functions";
import {
  DEST_FILE_SRC,
  END_ID,
  END_SHIFT,
  MAIN_FILE_SRC,
  MODIFIED_SRC,
  START_ID,
  START_SHIFT,
} from "./configs";

async function main(
  src: string,
  startId: string,
  endId: string,
  destSrc: string,
  modifiedSrc: string,
  startShift: number,
  endShift: number
) {
  importTimeSheet(
    exportTimeSheet(readXLS(src), startId, endId, startShift, endShift),
    destSrc,
    modifiedSrc
  );
}

main(
  MAIN_FILE_SRC,
  START_ID,
  END_ID,
  DEST_FILE_SRC,
  MODIFIED_SRC,
  START_SHIFT,
  END_SHIFT
);
