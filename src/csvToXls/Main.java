package csvToXls;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.FileReader;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.opencsv.CSVReader;

public class Main {

	public static void main(String[] args) throws IOException {
		// TODO Auto-generated method stub
		CSVReader reader = new CSVReader(new FileReader("emps.csv"), ',');
		List<Employee> emps = new ArrayList<Employee>();

		List<String[]> records = reader.readAll();

		Iterator<String[]> iterator = records.iterator();

		while (iterator.hasNext()) {
			String[] record = iterator.next();
			Employee emp = new Employee();
			emp.setId(record[0]);
			emp.setName(record[1]);
			emp.setAge(record[2]);
			emp.setCountry(record[3]);
			emps.add(emp);
		}

		System.out.println(emps);


		
		reader.close();
		
		
		
		//poi
		Workbook wb = new HSSFWorkbook();
	    //Workbook wb = new XSSFWorkbook();
	    CreationHelper createHelper = wb.getCreationHelper();
	    Sheet sheet = wb.createSheet("new sheet");

	    
	    Iterator<Employee> iterator2 = emps.iterator();
	    
	    int rowNumber = 0;
		while (iterator2.hasNext()) {
			Employee emp = iterator2.next();
			
			
//			System.out.println(emp.getName());
//			emp.getId();
//			emp.getName();
//			emp.getAge();
//			emp.getCountry();
			
		    // Create a row and put some cells in it. Rows are 0 based.
		    Row row = sheet.createRow((short)rowNumber);

		    // Create a cell and put a value in it.
		    Cell cell = row.createCell(0);
		    cell.setCellValue(emp.getId());
		    
		    // Or do it on one line.  1 2 and 3 are cell number
		    row.createCell(1).setCellValue(emp.getName());
		    row.createCell(2).setCellValue(
		         createHelper.createRichTextString(emp.getAge()));
		    row.createCell(3).setCellValue(emp.getCountry());
		    
		    
		   rowNumber = rowNumber+1;
		}
		
	    // Write the output to a file
	    FileOutputStream fileOut = new FileOutputStream("workbook.xls");
	    wb.write(fileOut);
	    fileOut.close();
	    System.out.println("write successful");
	}

}
