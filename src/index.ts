import { exportTimeSheet, importTimeSheet, readXLS } from "./functions";
import {
  DEST_FILE_SRC,
  END_ID,
  MAIN_FILE_SRC,
  MODIFIED_SRC,
  START_ID,
} from "./configs";

async function main(
  src: string,
  startId: string,
  endId: string,
  destSrc: string,
  modifiedSrc: string
) {
  importTimeSheet(
    exportTimeSheet(readXLS(src), startId, endId),
    destSrc,
    modifiedSrc
  );
}

main(MAIN_FILE_SRC, START_ID, END_ID, DEST_FILE_SRC, MODIFIED_SRC);
