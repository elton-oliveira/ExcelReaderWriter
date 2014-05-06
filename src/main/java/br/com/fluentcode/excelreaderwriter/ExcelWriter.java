package br.com.fluentcode.excelreaderwriter;

import java.math.BigDecimal;
import java.util.Date;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

/**
 * 
 * Abstract component for writing excel document.
 * Must implement the method {@link #createWorkbook()}
 *
 */
public abstract class ExcelWriter {

	private static final String DEFAULT_SHEET_NAME = "Plan1";
	private static final String DOUBLE_CELL_STYLE = "#.##";
	private static final String DATE_CELL_STYLE = "dd/MM/yyyy";

	/**
	 * Create Workbook: HSSFWorkbook for xls and XSSFWorkbook for xlsx
	 *
	 */
	protected abstract Workbook createWorkbook();
	
	/**
	 * Invoked after set the cell value, can be used to set formatting.
	 * Its implementation is not mandatory.
	 */
	protected void postCellValue(Cell cell){	}

	/**
	 * Writes excel  with the default name spreadsheet: Plan1
	 * 
	 * @param rowList the data to compose the spreadsheet
	 * @return Workbook the representation excel document
	 */
	public <T> Workbook writeExcel(List<T[]> rowList) {
		return escreverExcel(DEFAULT_SHEET_NAME, rowList);
	}

	/**
	 * Write the excel document
	 * 
	 * @param sheetName the sheet name
	 * @param rowList the data to compose the spreadsheet
	 * @return Workbook the representation excel document
	 */
	public <T> Workbook escreverExcel(String sheetName, List<T[]> rowList) {
		Workbook workbook = createWorkbook();
		Sheet sheet = workbook.createSheet(sheetName);
		CellStyle doubleCellStyle = workbook.createCellStyle();
		doubleCellStyle.setDataFormat(workbook.getCreationHelper().createDataFormat()
				.getFormat(DOUBLE_CELL_STYLE));
		CellStyle dateCellStyle = workbook.createCellStyle();
		dateCellStyle.setDataFormat(workbook.getCreationHelper().createDataFormat()
				.getFormat(DATE_CELL_STYLE));
		for (int i = 0; i < rowList.size(); i++) {
			T[] rowValues = rowList.get(i);
			Row row = sheet.createRow(i);
			for (int j = 0; j < rowValues.length; j++) {
				T value = rowValues[j];
				Cell cell = row.createCell(j);
				if (value instanceof String) {
					cell.setCellValue((String) value);
				} else if (value instanceof Double) {
					cell.setCellStyle(doubleCellStyle);
					cell.setCellValue((Double) value);
				} else if (value instanceof Integer) {
					cell.setCellValue((Integer) value);
				} else if (value instanceof Date) {
					cell.setCellStyle(dateCellStyle);
					cell.setCellValue((Date) value);
				} else if (value instanceof Boolean) {
					cell.setCellValue(String.valueOf(value));
				} else if (value instanceof BigDecimal) {
					cell.setCellValue(Double.valueOf(value.toString()));
					cell.setCellStyle(doubleCellStyle);
				} else if (value != null) {
					cell.setCellValue(value.toString());
				}
				
				this.postCellValue(cell);
			}
		}
		return workbook;
	}

}
