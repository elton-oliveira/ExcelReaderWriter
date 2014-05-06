package br.com.fluentcode.excelreaderwriter;

import java.io.ByteArrayOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Workbook;

public class ExcelBinaryGenerator {

	public byte[] generateByteArray(Workbook workbook) throws IOException {
		ByteArrayOutputStream bos = new ByteArrayOutputStream();
		workbook.write(bos);
		return  bos.toByteArray();
	}
}
