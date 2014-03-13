package br.com.fluentcode.poi;

import java.io.ByteArrayOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Workbook;

public class ExcelGeradorBinario {

	public byte[] gerarByteArray(Workbook workbook) throws IOException {
		ByteArrayOutputStream bos = new ByteArrayOutputStream();
		workbook.write(bos);
		return  bos.toByteArray();
	}
}
