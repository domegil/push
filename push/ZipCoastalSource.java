
import java.io.FileInputStream;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


public class ZipCoastalSource {
	public List<ZipCoastal> excelImport(){
		List<ZipCoastal> listZipCoastal=new ArrayList<>();
		long sid=0;
		String szip="";
		String scounty="";
		Boolean scoastal="";
		
		String excelFilePath=".\\Coastal.xlsx";
		
		long start = System.currentTimeMillis();
		
		FileInputStream inputStream;
		try {
			inputStream = new FileInputStream(excelFilePath);
			Workbook workbook=new XSSFWorkbook(inputStream);
			Sheet firstSheet=workbook.getSheetAt(0);
			Iterator<Row> rowIterator=firstSheet.iterator();
			rowIterator.next();
			
			while(rowIterator.hasNext()) {
				Row nextRow = rowIterator.next();
				Iterator<Cell> cellIterator=nextRow.cellIterator();
				while(cellIterator.hasNext()) {
					Cell nextCell=cellIterator.next();
					int columnIndex=nextCell.getColumnIndex();
					switch (columnIndex) {
					case 0:
						sid=(long)nextCell.getNumericCellValue();
						System.out.println(sid);
						break;
					case 1:
						szip=nextCell.getStringCellValue();
						System.out.println(szip);
						break;
					case 2:
						scounty=nextCell.getStringCellValue();
						System.out.println(scounty);
						break;
					case 3:
						scoastal=nextCell.getBooleanCellValue();
						System.out.println(scoastal);
						break;					
					}
					listZipCoastal.add(new ZipCoastal(sid, szip, scounty, scoastal));
				}
			}
			
			workbook.close();
			long end = System.currentTimeMillis();
			System.out.printf("Import done in %d ms\n", (end - start));
			
		} catch (Exception e) {
			// TODO Auto-generated catch block
			
			e.printStackTrace();
		}
		
		return listZipCoastal;
	}

}
