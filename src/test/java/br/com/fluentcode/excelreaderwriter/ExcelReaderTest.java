package br.com.fluentcode.excelreaderwriter;

import static org.hamcrest.CoreMatchers.equalTo;
import static org.hamcrest.MatcherAssert.assertThat;

import java.io.ByteArrayInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.math.BigDecimal;
import java.util.ArrayList;
import java.util.GregorianCalendar;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Workbook;
import org.junit.Before;
import org.junit.Test;

public class ExcelReaderTest {

	private ExcelReader reader;
	private InputStream stream;

	@Before
	public void setup() throws IOException {
		
		reader = new ExcelReader();
		
		ExcelWriter writer = new ExcelWriter() {
			@Override
			protected Workbook createWorkbook() {
				return new HSSFWorkbook() ;
			}
		};
		
		List<Object[]> rowList = new ArrayList<Object[]>();
		Object[] rowValues0 = {2, 2000.568, Integer.MAX_VALUE, new BigDecimal("78776666.789")};
		Object[] rowValues1 = {new GregorianCalendar(2014, 02, 13).getTime(), true};
		rowList.add(rowValues0);
		rowList.add(rowValues1);
		
		Workbook workbook = writer.writeExcel(rowList);
		
		//Gets the workbook input stream
		ExcelBinaryGenerator binario = new ExcelBinaryGenerator();
		byte[] byteArray = binario.generateByteArray(workbook);
		stream = new ByteArrayInputStream(byteArray);
	}

	@Test
	public void shouldReadExcel() {
		List<String[]> rows = reader.readExcel(stream);
		String[] row0 = rows.get(0);
		String[] row1 = rows.get(1);
		assertThat(row0[0], equalTo("2"));
		assertThat(row0[1], equalTo("2000.568"));
		assertThat(row0[2], equalTo(String.valueOf(Integer.MAX_VALUE)));
		assertThat(row0[3], equalTo("78776666.789"));
		assertThat(row1[0], equalTo("13/03/2014"));
		assertThat(row1[1], equalTo("true"));
	}
	
}
