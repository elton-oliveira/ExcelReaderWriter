package br.com.fluentcode.poi;

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
 * Componente abstrato para escrita de excel.
 * Deve obrigatoriamente implementar o método {@link #criarWorkbook()}
 *
 */
public abstract class ExcelEscritor {

	private static String DEFAULT_SHEET_NAME = "Plan1";
	private static String DOUBLE_CELL_STYLE = "#.##";
	private static String DATE_CELL_STYLE = "dd/MM/yyyy";

	/**
	 * Cria um Workbook: HSSFWorkbook para xls e XSSFWorkbook para xlsx
	 *
	 */
	protected abstract Workbook criarWorkbook();
	
	/**
	 * 
	 * Invocado após setar o valor da célula, pode ser utilizado para setar uma formatação
	 * Sua implementação não é obrigatória
	 */
	protected void postCellValue(Cell cell){	}

	/**
	 * Escreve o excel com nome padrão de planilha: Plan1
	 * 
	 * @param rowList dados para compor a planilha
	 * @return Workbook representação da planilha
	 */
	public <T> Workbook escreverExcel(List<T[]> rowList) {
		return escreverExcel(DEFAULT_SHEET_NAME, rowList);
	}

	/**
	 * Escreve o excel
	 * 
	 * @param sheetName nome da planilha
	 * @param rowList dados para compor a planilha
	 * @return Workbook representação da planilha
	 */
	public <T> Workbook escreverExcel(String sheetName, List<T[]> rowList) {
		Workbook workbook = criarWorkbook();
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
