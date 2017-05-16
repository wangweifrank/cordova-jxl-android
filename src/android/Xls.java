package cordova.jxl.android;

import java.io.File;
import java.io.IOException;
import java.util.Iterator;

import org.apache.cordova.CallbackContext;
import org.apache.cordova.CordovaPlugin;
import org.json.JSONArray;
import org.json.JSONException;
import org.json.JSONObject;

import android.os.Environment;
import android.util.Log;
import jxl.Sheet;
import jxl.Workbook;
import jxl.format.Alignment;
import jxl.read.biff.BiffException;
import jxl.write.Label;
import jxl.write.WritableCellFormat;
import jxl.write.WritableFont;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;
import jxl.write.biff.RowsExceededException;

public class Xls extends CordovaPlugin {
	public static final String ACTION_SAVE_XLS = "saveXLS";

	private static final String TAG = "xls";

	public String dirname;

	@Override
	public boolean execute(String action, JSONArray args, CallbackContext callbackContext) throws JSONException {
		Log.v(TAG, "Xls CordovaPlugin execute" + action);
		try {
			if (ACTION_SAVE_XLS.equals(action)) {

				JSONObject params = args.getJSONObject(0);

				try {
					// 存储路径
					this.dirname = params.getString("dirname");

					// 创建Excel文件
					WritableWorkbook wb = this.createWorkbook(params.getString("filename"));

					// 创建Sheet
					WritableSheet sheetObject = this.createSheet(wb, params.getString("sheetdata"), 3);

					// 写入Sheet 数据
					JSONArray lineItems = params.getJSONArray("data");
					writeData(sheetObject, lineItems);

					// 写入到文件中
					wb.write();
					// 关闭文件
					wb.close();

					// 写入成功回调JS
					callbackContext.success(params);
					return true;
				} catch (IOException e) {
					Log.e(TAG, "write xls error", e);
				}
			}

			callbackContext.error("Invalid action");
			return false;
		} catch (Exception e) {
			System.err.println("Exception: " + e.getMessage());
			callbackContext.error(e.getMessage());
			Log.e(TAG, "execute xls cordova plugin error ", e);
			return false;
		}
	}

	// 每行数据加载到cell中
	public void writeData(WritableSheet wsheet, JSONArray lineItems) throws JSONException {
		int rowIndex = wsheet.getRows();
		if (rowIndex == 0) {
			Log.v(TAG, "need to write title index");
			// 写入title
			jsonObjectToCell(wsheet, lineItems.getJSONObject(0), 0, true);
			rowIndex++;
		}
		for (int i = 0, size = lineItems.length(); i < size; i++) {
			jsonObjectToCell(wsheet, lineItems.getJSONObject(i), (rowIndex + i), false);
		}
	}

	public void jsonObjectToCell(WritableSheet sheetObj, JSONObject obj, int rowPosition, boolean isTitle) {
		try {
			int columnPosition = 0;
			Iterator<?> keys = obj.keys();
			while (keys.hasNext()) {
				// 获取title key
				String key = (String) keys.next();

				// 获取key对应value
				String value = obj.getString(key);

				// 判断是否写入title
				if (isTitle) {
					// 写入title数据
					this.writeCell(columnPosition, rowPosition, key, true, sheetObj);
				} else {
					// 写入单元数据
					this.writeCell(columnPosition, rowPosition, value, false, sheetObj);
				}
				// 写入列指针自增1
				columnPosition++;
			}
			Log.i(TAG, "write data rowPosition:" + rowPosition + " finished!");
		} catch (JSONException e) {
			Log.e(TAG, "[ERROR] parse data JsonObject Error ", e);
		} catch (WriteException e) {
			Log.e(TAG, "[ERROR] write data cell error", e);
		}
	}

	/**
	 * @param fileName
	 *            - the name to give the new workbook file
	 * @return - a new WritableWorkbook with the given fileName
	 * @throws BiffException
	 */
	public WritableWorkbook createWorkbook(String fileName) throws BiffException {
		Log.v(TAG, "create Excel file fileName:" + fileName);

		// get the sdcard's directory
		File sdCard = Environment.getExternalStorageDirectory();
		// add on the your app's path
		File dir = new File(sdCard.getAbsolutePath() + "/" + this.dirname);
		if (!dir.exists()) {
			// make them in case they're not there
			dir.mkdirs();
		}
		WritableWorkbook wb = null;
		try {
			// create a standard java.io.File object for the Workbook to use
			File wbfile = new File(dir, fileName);
			if (!wbfile.exists()) {
				wbfile.createNewFile();
				// create a new WritableWorkbook using the java.io.File and
				// WorkbookSettings from above
				wb = Workbook.createWorkbook(wbfile);
			} else {
				wb = Workbook.createWorkbook(wbfile, Workbook.getWorkbook(wbfile));
			}
		} catch (IOException e) {
			Log.e(TAG, "[ERROR] create Excel File exception!!", e);
		}
		return wb;
	}

	/**
	 * @param wb
	 *            - WritableWorkbook to create new sheet in
	 * @param sheetName
	 *            - name to be given to new sheet
	 * @param sheetIndex
	 *            - position in sheet tabs at bottom of workbook
	 * @return - a new WritableSheet in given WritableWorkbook
	 */
	public WritableSheet createSheet(WritableWorkbook wb, String sheetName, int sheetIndex) {
		int sheetNumb = wb.getNumberOfSheets();

		int index = sheetIndex;

		for (index = 0; index < sheetNumb; index++) {
			Sheet sheet = wb.getSheet(index);
			if (sheet.getName().equals(sheetName)) {
				Log.v(TAG, "create sheet sheetName:" + sheetName + " index:" + index);
				return wb.getSheet(index);
			}
		}
		Log.v(TAG, "create sheet sheetName:" + sheetName + " index:" + index);
		// create a new WritableSheet and return it
		return wb.createSheet(sheetName, index);
	}

	/**
	 * @param columnPosition
	 *            - column to place new cell in
	 * @param rowPosition
	 *            - row to place new cell in
	 * @param contents
	 *            - string value to place in cell
	 * @param headerCell
	 *            - whether to give this cell special formatting
	 * @param sheet
	 *            - WritableSheet to place cell in
	 * @throws RowsExceededException
	 *             - thrown if adding cell exceeds .xls row limit
	 * @throws WriteException
	 *             - Idunno, might be thrown
	 */
	public void writeCell(int columnPosition, int rowPosition, String contents, boolean headerCell, WritableSheet sheet)
			throws RowsExceededException, WriteException {
		// create a new cell with contents at position
		Label newCell = new Label(columnPosition, rowPosition, contents);

		if (headerCell) {
			// give header cells size 10 Arial bolded
			WritableFont headerFont = new WritableFont(WritableFont.ARIAL, 10, WritableFont.BOLD);
			WritableCellFormat headerFormat = new WritableCellFormat(headerFont);
			// center align the cells' contents
			headerFormat.setAlignment(Alignment.CENTRE);
			newCell.setCellFormat(headerFormat);
		}

		sheet.addCell(newCell);
	}
}
