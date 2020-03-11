package TestPackage;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.Map;
import java.util.Set;
import java.util.TreeMap;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class CompareExcel3 {

	public static void main(String[] args) {

		try {
			FileInputStream f1 = new FileInputStream("C:\\Users\\user\\Desktop\\TestBook1.xlsx");
			XSSFWorkbook wb1 = new XSSFWorkbook(f1);
			XSSFSheet sh1 = wb1.getSheetAt(0);
			int rowCount1 = sh1.getPhysicalNumberOfRows();

			FileInputStream f2 = new FileInputStream("C:\\Users\\user\\Desktop\\TestBook3.xlsx");
			XSSFWorkbook wb2 = new XSSFWorkbook(f2);
			XSSFSheet sh2 = wb2.getSheetAt(0);
			int rowCount2 = sh2.getPhysicalNumberOfRows();

			XSSFWorkbook wb3 = new XSSFWorkbook();
			XSSFSheet sh3 = wb3.createSheet("MissingDataSheet");
			XSSFRow writeRow;
			int writeRowCount = 1;
			Map<String, Object[]> missingData = new TreeMap<String, Object[]>();
			missingData.put("0", new Object[]{"Original", "Duplicate"});
			
			
			// compare row number for both sheet
			if (rowCount1 == rowCount2) {
				
				// For loop for row staring from first row to last row
				for (int i = 1; i < rowCount1; i++) {
					
					// getting total column number for specific row for both sheet 
					int colCount1 = sh1.getRow(i).getLastCellNum();
					int colCount2 = sh2.getRow(i).getLastCellNum();
					
					// compare row number for both sheet
					if (colCount1 == colCount2) {
						
						// For loop for column staring from first column to last column
						for (int j = 1; j < colCount2; j++) {
							XSSFRow row1 = sh1.getRow(i);
							XSSFRow row2 = sh2.getRow(i);

							// Getting cell value for sheet1
							String sh1_column1 = "";
							XSSFCell sh1_col1 = row1.getCell(j);
							if (sh1_col1 != null) {
								sh1_column1 = sh1_col1.getStringCellValue();
							}

							// Getting cell value for sheet2
							String sh2_column1 = "";
							XSSFCell sh2_col1 = row2.getCell(j);
							if (sh2_col1 != null) {
								sh2_column1 = sh2_col1.getStringCellValue();
							}

							if (sh1_column1.equals(sh2_column1)) {
								System.out.println(
										"Row : " + i + " Column : " + j + ", value is matched as--> " + sh1_column1);
							} else {
								System.out.println(
										"Row : " + i + " Column : " + j + ", value is not matched. Sheet1 value--> "
												+ sh1_column1 + " and Sheet2 value--> " + sh2_column1);
								
								String r = String.valueOf(writeRowCount);
								
							// if not equal then store in map
								missingData.put(r, new Object[]{sh1_column1, sh2_column1});
								writeRowCount++;
							}
						}
					}else{
						System.out.println("Column number is not equal for both Sheet. Sheet1 column count : " + rowCount1
								+ " and Sheet2 column count : " + rowCount2);
					}
				}

				// To write missing data in excel sheet
				Set <String> keyId = missingData.keySet();
				int rowid = 0;
				
				for(String key : keyId){
					writeRow = sh3.createRow(rowid++);
					Object[] objectArr = missingData.get(key);
					int cellId =0;
					
					for(Object obj : objectArr){
						Cell cell = writeRow.createCell(cellId++);
						cell.setCellValue((String)obj);
					}
				}
				
				// write the workbook
				FileOutputStream out = new FileOutputStream(new File("C:\\Users\\user\\Desktop\\MissingDataExcel.xlsx"));
				wb3.write(out);
				out.close();
				System.out.println("Written successfully");
				
			} else {
				System.out.println("Row number is not equal for both Sheet. Sheet1 row count : " + rowCount1
						+ " and SHeet2 row count : " + rowCount2);
			}

		} catch (Exception e) {
			e.printStackTrace();
		}

	}
}