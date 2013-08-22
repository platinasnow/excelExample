import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.xmlbeans.impl.xb.ltgfmt.TestsDocument.Tests;


public class Main {

	/**
	 * @param args
	 */
	public static void main(String[] args) {

		Object[][] excelTable;
		ExcelDao excel = new ExcelDao();
		try {
			excelTable =  excel.loadExcel();
			
			excel.excelWrite();
			
			
			String result ="";
			String testString = "aaaa fad f ba b ewa sadaf [[ㅁㅁㅁ]]abbad wtw241vapsd [[ㅁ5ㅁ]] vaaasdfa[[주소|이름]] bhhhhhhhhhhhhhhh";
			
			while(testString.indexOf("]]") !=-1){
				if(testString.indexOf("[[") != 0){
					result += testString.substring(0,testString.indexOf("[["));
					String realText = testString.substring(testString.indexOf("[[")+2,testString.indexOf("]]"));
					if(realText.indexOf("|") != -1){
						result += "<a href='"+realText.substring(0,realText.indexOf("|"))+"'>"+realText.substring(realText.indexOf("|")+1,realText.length())+"</a>";
						testString = testString.substring(testString.indexOf("]]")+2,testString.length());
					}else{
						result += "<a href='"+realText+"'>"+realText+"</a>";
						testString = testString.substring(testString.indexOf("]]")+2,testString.length());
					}
					
				}
			}
			System.out.println(testString.indexOf("[["));
			System.out.println("string :"+testString);
			System.out.println("result :" + result);
			
			
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}
		
	}

}
