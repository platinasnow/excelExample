import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;


public class ExcelDao {

	public Object[][] loadExcel() throws FileNotFoundException, IOException{
		File file = new File("C:\\Users\\Administrator\\Documents\\ex01.xls");
		POIFSFileSystem fs = new POIFSFileSystem(new FileInputStream(file));
		HSSFWorkbook wb = new HSSFWorkbook(fs);
		
		 HSSFSheet sheet = wb.getSheetAt(0);
		 int rows = sheet.getPhysicalNumberOfRows();
		 int cells  = sheet.getRow(0).getPhysicalNumberOfCells();
		 System.out.println("row :"+rows);
		 System.out.println("cells :"+ cells);
		 Object[][] table = new Object[rows][cells];
		 
		 for(int rowIdx = 0; rowIdx < rows ; rowIdx++){
			 HSSFRow row = sheet.getRow(rowIdx);
			 if(row != null  ){
				 for(int cellIdx = 0; cellIdx < cells ; cellIdx++){
					 HSSFCell cell = row.getCell(cellIdx);
					 if(cell !=null ){
						 switch(cell.getCellType()){
							 case HSSFCell.CELL_TYPE_STRING:
								 table[rowIdx][cellIdx] = cell.getStringCellValue();
								 break;
							 case HSSFCell.CELL_TYPE_BOOLEAN:
								 table[rowIdx][cellIdx] = cell.getBooleanCellValue();
								 break;
							 case HSSFCell.CELL_TYPE_NUMERIC:
								 table[rowIdx][cellIdx] = cell.getNumericCellValue();
								 break;
							 case HSSFCell.CELL_TYPE_BLANK:
								 table[rowIdx][cellIdx] = "";
								 break;
							 case HSSFCell.CELL_TYPE_FORMULA:
								 table[rowIdx][cellIdx] = cell.getCellFormula(); 
								 break;
						 }
					 }
				 }
			 }
		 }
		 return table;
	}
	
	public void excelWrite(){
		HSSFWorkbook workbook = new HSSFWorkbook();
		HSSFSheet sheet = workbook.createSheet();
		
		HSSFFont font = workbook.createFont();
		font.setFontName(HSSFFont.FONT_ARIAL);
		
		HSSFCellStyle titleStyle = workbook.createCellStyle();
		titleStyle.setFillBackgroundColor(HSSFColor.GREY_50_PERCENT.index);
		titleStyle.setFont(font);
		
		HSSFRow row = sheet.createRow(0);
		
		HSSFCell cell1 = row.createCell(0);
		cell1.setCellValue("力格1");
		cell1.setCellStyle(titleStyle);
		
		HSSFCell cell2 = row.createCell(1);
		cell2.setCellValue("力格2");
		cell2.setCellStyle(titleStyle);
		
		HSSFCell cell3 = row.createCell(2);
		cell3.setCellValue("力格3");
		cell3.setCellStyle(titleStyle);
		
		HSSFCellStyle contentsStyle = workbook.createCellStyle();
		contentsStyle.setFont(font);

		String realName = "Test.xls";
		File file = new File("C:\\Users\\Administrator\\Documents\\"+realName);
		
		try {
			FileOutputStream fos = new FileOutputStream(file);
			workbook.write(fos);
			fos.close();
		} catch (IOException e) {
			e.printStackTrace();
		}
		
			
		
	}
	
}
