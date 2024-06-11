import java.io.File;  
import java.util.Scanner;
import java.io.FileInputStream;  
import java.io.IOException;  
import org.apache.poi.xssf.usermodel.XSSFSheet;  
import org.apache.poi.xssf.usermodel.XSSFWorkbook;  
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.FormulaEvaluator;  
import org.apache.poi.ss.usermodel.Row;  
import java.text.SimpleDateFormat;
public class ReadBasicExcel  
{  
	public static void main(String args[]) throws IOException  
	{  
		try{
		String InputString,TableName,ExelPath,ColumnString="";
		System.out.println("Enter the Table Name and the path (TableName~Path) ");
		Scanner SC = new Scanner(System.in);
		InputString = SC.nextLine();
		SC.close();
		
		TableName = InputString.substring(0,InputString.indexOf("~"));
		ExelPath=InputString.substring(InputString.indexOf("~")+1,InputString.length());
		System.out.println(TableName);
		System.out.println(ExelPath);
		
		FileInputStream fis=new FileInputStream(new File("D:\\Databook.xlsx"));  
		@SuppressWarnings("resource")
		XSSFWorkbook wb=new XSSFWorkbook(fis);   
		XSSFSheet sheet=wb.getSheetAt(0);  
		
		Row intialRow = sheet.getRow(0);
		
		for (Cell cell :  intialRow) {
			ColumnString+=cell.getStringCellValue()+",";
		}
		
		StringBuffer ColumnNameBuffer= new StringBuffer(ColumnString);  
		ColumnNameBuffer.deleteCharAt(ColumnNameBuffer.length()-1);   
		ColumnString = ColumnNameBuffer.toString();
		ColumnString = ColumnString.replaceAll(" ", "");

		
		int rowCount=1;
		String columnData="";
		while(sheet.getRow(rowCount) != null) {
			int stringCol = 0 , intCol=0 , blankCol=0 , defCount=0;
			for(Cell cellData : sheet.getRow(rowCount)) {
			switch(cellData.getCellTypeEnum())  
				{  

					case NUMERIC: 
											if(cellData.getNumericCellValue()==0.0) {
											columnData += "'NULL' ,"; }
											else if(DateUtil.isCellDateFormatted(cellData)){
											columnData += cellData.getDateCellValue() +",";}
											else {columnData += +cellData.getNumericCellValue()+" ,"; }
											intCol =intCol+1;
											break; 
					case STRING: 
											if(cellData.getStringCellValue().trim().isEmpty())
											{columnData += "'NULL' ,";}
											else
											{columnData += "'"+cellData.getStringCellValue()+"' ,";} stringCol=stringCol+1;break; 
					
					case BLANK:          	blankCol=blankCol+1;break;
					
					default : 				defCount +=1;
					
				}
				
				//if (cellData==null) { System.out.print("Blank / Null Spotted ");}else {System.out.println("not Blank");}
				
			}
			
			System.out.println("Def Count  "+ Integer.toString(defCount)  + " Blank count  "+ Integer.toString(blankCol)+ " Int Count  " + Integer.toString(intCol)+"String Count  "+ Integer.toString(stringCol)+" Total :- "+Integer.toString(blankCol+intCol+stringCol));
			
			StringBuffer ColumnDataBuffer= new StringBuffer(columnData);  
			ColumnDataBuffer.deleteCharAt(ColumnDataBuffer.length()-1);   
			columnData =ColumnDataBuffer.toString();
			columnData = columnData.replaceAll(" ", "");
			
			//System.out.println("Insert into " + TableName + "("+ColumnString+") VALUES (" + columnData+")");
			rowCount =rowCount+1;
			columnData = "";
		}
		
		
		
		}catch(Exception error) {
			System.out.println("The entered tablename and path are not in correct format.");
			System.out.println(error.getMessage());
		}
		
	}  
}  