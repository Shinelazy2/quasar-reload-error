import { ipcMain as im } from 'electron';
import { commonExcel } from './common-excel';
import txtToDb from './txt-to-db';

im.handle('readTxtFile', async (_, txtFilePath: string) => {
  return await txtToDb(txtFilePath);
});
