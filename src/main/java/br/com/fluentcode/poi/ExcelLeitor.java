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
 * Componente para leitura de documentos excel.
 *
 */
public class ExcelLeitor {
	
	private SimpleDateFormat dateFormat;
	
	/**
	 * 
	 * @param dateFormat formato que deve ser retornado os valores das células do tipo data.
	 * Se for utilizado o construtor sem argumentos será utilizado o formato 'dd/MM/yyyy'
	 */
	public ExcelLeitor(SimpleDateFormat dateFormat) {
		this.dateFormat = dateFormat;
	}
	
	public ExcelLeitor() {
		dateFormat = new SimpleDateFormat("dd/MM/yyyy");
	}

	/**
	 * Realiza a leitura da primeira sheet do excel
	 * 
	 * @param stream representa o input stream do excel
	 * @return o resultado da letura onde cada linha do excel é armazenado em um array de String
	 */
	public List<String[]> lerExcel(InputStream stream) {
		return lerExcel(stream, 0);
	}

	/**
	 * 
	 * @param stream representa o input stream do excel
	 * @param sheetIndex representa o index (baseado em 0) da sheet que deve ser lida
	 * @return o resultado da leitura onde cada linha do excel é armazenado em um array de String
	 */
	public List<String[]> lerExcel(InputStream stream, int sheetIndex) {
		Workbook workbook = this.createWorkbook(stream);
		Sheet sheet = workbook.getSheetAt(sheetIndex);
		if (sheet == null) {
			throw new IllegalArgumentException("Planilha inexistente");
		}
		FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();
		return read(sheet, evaluator);
	}

	/**
	 * 
	 * @param stream representa o input stream do excel
	 * @param sheetName representa o nome da sheet que deve ser lida
	 * @return o resultado da leitura onde cada linha do excel é armazenado em um array de String
	 */
	public List<String[]> lerExcel(InputStream stream, String sheetName) {
		Workbook workbook = this.createWorkbook(stream);
		Sheet sheet = workbook.getSheet(sheetName);
		if (sheet == null) {
			throw new IllegalArgumentException("Planilha inexistente");
		}
		FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();
		return read(sheet, evaluator);
	}

	/*
	 * Cria o workbook, a representação do excel
	 */
	private Workbook createWorkbook(InputStream stream) {
		Workbook workbook = null;
		try {
			workbook = WorkbookFactory.create(stream);
		} catch (Exception e) {
			throw new IllegalArgumentException("Erro ao tentar criar o excel", e);
		}
		return workbook;
	}

	/*
	 * Realiza a leitura da planilha
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
	 * Extrai e retorna o valor da célula em forma de String
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
