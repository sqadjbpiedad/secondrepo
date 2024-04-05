package testapp1;

import java.util.Set;
import java.util.TreeMap;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Iterator;
import java.util.Map;
import java.io.File;
import java.io.FileInputStream;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class PoiDemo {

	public static void main(String[] args) {
		
		//createworkbook("employees2", "records");
		//createworkbook("employees", "records");
		//readexcel("employees", "records");
		readexcel();
		//appendrow("employees", "records", "5", "heidi", "operations");
	}
	
	//append row
	public static void appendrow(String wb, String ws, String id, String name, String department) {
		
		try {
			//check if exists
			//FileInputStream file = new FileInputStream(wb + ".xlsx");
			File file = new File(wb + ".xlsx");
			if(file.exists()) {
				
				FileInputStream fileInputStream = new FileInputStream(file);
				
				XSSFWorkbook workbook = new XSSFWorkbook(fileInputStream);
				XSSFSheet sheet = workbook.getSheet(ws);
				int rowlastnum = sheet.getLastRowNum();
				Row newrow = sheet.createRow(rowlastnum + 1);
				
				Cell cell1 = newrow.createCell(0);
				cell1.setCellValue(id);
				//newrow.createCell(0).setCellValue(id); // equivalent of 2 line above
				
				Cell cell2 = newrow.createCell(1);
				cell2.setCellValue(name);
				
				Cell cell3 = newrow.createCell(2);
				cell3.setCellValue(department);
				
				//write to file
				FileOutputStream out = new FileOutputStream(file);
				workbook.write(out);
				System.out.println("new row added");
				out.close();
				
			} else {
				System.out.println("cannot append, file does not exist.");
			}
			
			
		} catch (Exception e) {
			e.printStackTrace();
		}
		
		
		
	}
	
	
	
	
	
	
	public static void readexcel() {
		
		try{
			FileInputStream file = new FileInputStream("employees.xlsx");
			XSSFWorkbook workbook = new XSSFWorkbook(file);
			XSSFSheet sheet  = workbook.getSheet("records");
			//XSSFSheet sheet  = workbook.getSheetAt(0);
			
			//loop over rows in sheet
			Iterator<Row> rowiterator = sheet.rowIterator();
			while(rowiterator.hasNext()) {
				Row row = rowiterator.next();
				
				//loop over columns in each row
				Iterator<Cell> celliterator = row.cellIterator();
				while(celliterator.hasNext()) {
					Cell cell = celliterator.next();
					System.out.print(cell.getStringCellValue() + "\t");
				} //end of column loop
				System.out.print("\n");
			} //end of row loop
			file.close();
			System.out.println("-----end-----");
			
		} catch (IOException e){ //InvalidFormatException 
			//custom error
			System.out.println("file not found");
		}
	}
	

	public static void readexcel(String workbookname, String worksheetname) {
		
		try{
			FileInputStream file = new FileInputStream(workbookname + ".xlsx");
			XSSFWorkbook workbook = new XSSFWorkbook(file);
			XSSFSheet sheet  = workbook.getSheet(worksheetname);
			//XSSFSheet sheet  = workbook.getSheetAt(0);
			
			//loop over rows in sheet
			Iterator<Row> rowiterator = sheet.rowIterator();
			while(rowiterator.hasNext()) {
				Row row = rowiterator.next();
				
				//loop over columns in each row
				Iterator<Cell> celliterator = row.cellIterator();
				while(celliterator.hasNext()) {
					Cell cell = celliterator.next();
					System.out.print(cell.getStringCellValue() + "\t");
				} //end of column loop
				System.out.print("\n");
			} //end of row loop
			file.close();
			System.out.println("-----end-----");
			
		} catch (IOException e){
			//custom error
			System.out.println("file not found");
		}
	}
	
	
	public static void createworkbook() {
		//write to xlsx
		//create instance workbook	
		XSSFWorkbook workbook = new XSSFWorkbook();
		XSSFSheet sheet = workbook.createSheet("Employees");
		XSSFSheet sheet1 = workbook.createSheet("Employees1");
		XSSFSheet sheet2 = workbook.createSheet("Employees2");
		
		//data
		Map<String,Object[]> data = new TreeMap<String, Object[]>();
		data.put("1", new Object[] {"id", "name", "department"});
		data.put("2", new Object[] {"1", "joseph", "mis"});
		data.put("3", new Object[] {"2", "ryan", "hr"});
		data.put("4", new Object[] {"3", "didi", "accounting"});
		
		Set<String> keyset = data.keySet();
		
		int rownum = 0;
		
		//loop
		for(String key:keyset) {
			
			Row row = sheet.createRow(rownum++);
			Object[] obj = data.get(key);
			
			int cellnum = 0;
			//for each column in each row
			for(Object o:obj) {
				Cell cell = row.createCell(cellnum++);
				cell.setCellValue(o.toString());
				
			}			
		}
		
		//write file in filesystem
		try {
			FileOutputStream out = new FileOutputStream(new File("employees.xlsx"));
			workbook.write(out);
			out.close();
			System.out.println("write xlsx ok");
		} catch (Exception e) {
			System.out.println(e);
		}
	}
	
	
	public static void createworkbook(String workbookname, String worksheetname) {
		//write to xlsx
		//create instance workbook	
		XSSFWorkbook workbook = new XSSFWorkbook();
		XSSFSheet sheet = workbook.createSheet(worksheetname);
		
		//data
		Map<String,Object[]> data = new TreeMap<String, Object[]>();
		data.put("1", new Object[] {"id", "name", "department"});
		data.put("2", new Object[] {"1", "joseph", "mis"});
		data.put("3", new Object[] {"2", "ryan", "hr"});
		data.put("4", new Object[] {"3", "didi", "accounting"});
		
		Set<String> keyset = data.keySet();
		
		int rownum = 0;
		
		//loop
		for(String key:keyset) {
			
			Row row = sheet.createRow(rownum++);
			Object[] obj = data.get(key);
			
			int cellnum = 0;
			//for each column in each row
			for(Object o:obj) {
				Cell cell = row.createCell(cellnum++);
				cell.setCellValue(o.toString());
				
			}			
		}
		
		//write file in filesystem
		try {
			FileOutputStream out = new FileOutputStream(new File(workbookname + ".xlsx"));
			workbook.write(out);
			out.close();
			System.out.println("write xlsx ok");
		} catch (Exception e) {
			System.out.println(e);
		}
		
	}

}
