package webApplication.WriteExcelProgram;

import java.io.*;

import java.util.*;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class PrintingDataInExcel {

	@SuppressWarnings("resource")
	public static void main(String[] args) throws IOException {

	
		LinkedList<String[]> list = new LinkedList<>();
		BufferedReader bufferedReader = new BufferedReader(
				new FileReader(".\\Document\\raw_data_new.txt"));
		FileOutputStream fileOut = new FileOutputStream(".\\target\\workbook.xlsx");
		XSSFWorkbook wb = new XSSFWorkbook();
		XSSFSheet sheet = wb.createSheet("Data");
		XSSFRow row = null;
		String line = null;
		XSSFCell cell = null;

		while ((line = bufferedReader.readLine()) != null) {
		
			list.add(line.split("\\s+"));
			Pattern pat = Pattern.compile("\\s+");
			Matcher match = pat.matcher(line);

			while (match.find()) {
				match.replaceAll(",");
				int row_num = 0;
				for (String[] eachLine : list) {
					row = sheet.createRow(row_num++);
					int cell_num = 0;
					for (String value : eachLine) {
						cell = row.createCell(cell_num++);
						cell.setCellValue(value);
					}
				}
			}
		}
		System.out.println("written Successfully");
		wb.write(fileOut);
		fileOut.close();
	}
	
}
