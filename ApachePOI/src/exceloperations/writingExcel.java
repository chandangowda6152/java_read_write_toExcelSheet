package exceloperations;

import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class writingExcel {
	public static void main(String args[]) throws IOException {
		
	XSSFWorkbook workbook=new XSSFWorkbook();
	XSSFSheet sheet=workbook.createSheet("Student Details");
	
		Object studdata[][]= {	{"StudId","Name","Class"},
								{100,"David","First Standard"},
								{101,"Prem","Primary"},
								{102,"Latha","Second Standard"},
								{103,"Latha","Second Standard"},
						};
		
		
		int rows=studdata.length;
		int cols=studdata[0].length;
		
		System.out.println(rows); //4
		System.out.println(cols); //3
		
			for(int r=0;r<rows;r++) //0
			{
				XSSFRow row=sheet.createRow(r);
				
				for(int c=0;c<cols;c++) 
				{
					XSSFCell cell=row.createCell(c); //0
					Object value=studdata[r][c];
					
					if(value instanceof String)
						cell.setCellValue((String)value );
					if(value instanceof Integer)
						cell.setCellValue((Integer)value );
					if(value instanceof Boolean)
						cell.setCellValue((Boolean)value );
					
				}
			}
		String filePath=".\\datafiles\\student.xlsx";
		FileOutputStream outstream=new FileOutputStream(filePath);
		workbook.write(outstream);
		
		outstream.close();
		
		System.out.println("Student.zxls file written successfully......");
	}

}
