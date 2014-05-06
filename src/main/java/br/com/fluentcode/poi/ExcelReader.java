package br.com.fluentcode.poi;

import java.io.InputStream;
import java.math.BigDecimal;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellValue;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

/**
 * 
 * Component for reading excel document.
 *
 */
public class ExcelReader {
	
	private SimpleDateFormat dateFormat;
	
	/**
	 * 
	 * @param dateFormat the format that the cell values ​​of type date should be returned.
	 * If the no-argument constructor is used the format 'dd/MM/yyyy' will be used.
	 */
	public ExcelReader(SimpleDateFormat dateFormat) {
		this.dateFormat = dateFormat;
	}
	
	public ExcelReader() {
		dateFormat = new SimpleDateFormat("dd/MM/yyyy");
	}

	/**
	 * Read the first sheet
	 * 
	 * @param stream the input stream excel document
	 * @return the reading result where each line is stored in an array of String
	 */
	public List<String[]> readExcel(InputStream stream) {
		return readExcel(stream, 0);
	}

	/**
	 * Read the sheet whose index is passed as parameter.
	 * 
	 * @param stream the input stream excel document
	 * @param sheetIndex the index (0-based) of the sheet which is to be read
	 * @return the reading result where each line is stored in an array of String
	 */
	public List<String[]> readExcel(InputStream stream, int sheetIndex) {
		Workbook workbook = this.createWorkbook(stream);
		Sheet sheet = workbook.getSheetAt(sheetIndex);
		if (sheet == null) {
			throw new IllegalArgumentException("Absent sheet");
		}
		FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();
		return read(sheet, evaluator);
	}

	/**
	 * 
	 * @param stream the input stream excel document
	 * @param sheetName the sheet name which is to be read
	 * @return the reading result where each line is stored in an array of String
	 */
	public List<String[]> readExcel(InputStream stream, String sheetName) {
		Workbook workbook = this.createWorkbook(stream);
		Sheet sheet = workbook.getSheet(sheetName);
		if (sheet == null) {
			throw new IllegalArgumentException("Absent sheet");
		}
		FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();
		return read(sheet, evaluator);
	}

	/*
	 * Create the workbook
	 */
	private Workbook createWorkbook(InputStream stream) {
		Workbook workbook = null;
		try {
			workbook = WorkbookFactory.create(stream);
		} catch (Exception e) {
			throw new IllegalArgumentException("Error while trying to create workbook", e);
		}
		return workbook;
	}

	/*
	 * Read sheet
	 */
	private List<String[]> read(Sheet sheet, FormulaEvaluator evaluator) {
		List<String[]> rowList = new ArrayList<String[]>();
		for (Row row : sheet) {
			int rowSize = row.getLastCellNum();
			String[] rowValues = new String[rowSize];
			for (int cn = 0; cn < rowSize; cn++) {
				Cell cell = row.getCell(cn, Row.CREATE_NULL_AS_BLANK);
				String cellValue = this.getCellValue(cell, evaluator);
				rowValues[cn] = cellValue;
			}
			rowList.add(rowValues);
		}

		return rowList;
	}

	/*
	 * Extracts and returns the cell value as a String
	 */
	private String getCellValue(Cell cell, FormulaEvaluator evaluator) {
		String cellValue = null;
		switch (cell.getCellType()) {
			case Cell.CELL_TYPE_STRING:
				cellValue = cell.getStringCellValue();
				break;
			case Cell.CELL_TYPE_NUMERIC:
				if (DateUtil.isCellDateFormatted(cell)) {
					cellValue =  dateFormat.format(cell.getDateCellValue());
				} else {
					cell.setCellType(Cell.CELL_TYPE_STRING);
					cellValue = new BigDecimal(cell.getStringCellValue())
							.toPlainString();
				}
				break;
			case Cell.CELL_TYPE_BOOLEAN:
				cellValue = String.valueOf(cell.getBooleanCellValue());
				break;
			case Cell.CELL_TYPE_FORMULA:
					CellValue cv = evaluator.evaluate(cell);
					switch (cv.getCellType()) {
					case Cell.CELL_TYPE_BOOLEAN:
						cellValue = String.valueOf(cv.getBooleanValue());
						break;
					case Cell.CELL_TYPE_NUMERIC:
						cellValue = String.valueOf(cv.getNumberValue());
						break;
					case Cell.CELL_TYPE_STRING:
						cellValue = cv.getStringValue();
						break;
					}
					break;
		}
		return cellValue;
	}
	
}
