package copierExcel;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Iterator;
import java.nio.channels.FileChannel;
 
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFFormulaEvaluator;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.commons.io.FileUtils;
import org.apache.commons.io.IOUtils;

//https://www.codejava.net/coding/how-to-read-excel-files-in-java-using-apache-poi
public class MainCopier 
{
	
	public static void main(String []args) throws IOException
	{
		String directoryPath="Z:\\Shared Documents\\General\\009-Scripts\\001-FWS-2021-2022\\Students-info\\Admins only\\Spring 2022\\Idea\\Test\\Files\\";
		String timesheetTemplate = directoryPath+"C9-FWS-firstName-lastName-Timesheet.xlsx";
        String destinationFolder="Z:\\Shared Documents\\General\\009-Scripts\\001-FWS-2021-2022\\Students-info\\Admins only\\Spring 2022\\Idea\\Test\\Outputs\\"+"C9-FWS-firstName-lastName-Timesheet.xlsx";
        
        String studentInfo=directoryPath+"Student-Info.xlsx";
        
		File fileToCopy = new File(timesheetTemplate);
		
		 
		
		
		
		
		
		FileInputStream inputStream = new FileInputStream(new File(studentInfo));
         
        Workbook workbook = new XSSFWorkbook(inputStream);
        Sheet firstSheet = workbook.getSheetAt(0);
        Iterator<Row> iterator = firstSheet.iterator();
        
        while (iterator.hasNext()) 
        {
            Row nextRow = iterator.next();
            
            if(nextRow.getCell(0).getCellType()==CellType.NUMERIC)
            {

            	String firstName=nextRow.getCell(3).getStringCellValue();
            	String lastName=nextRow.getCell(4).getStringCellValue(); 
            	String email=nextRow.getCell(5).getStringCellValue();
            	double empID=nextRow.getCell(0).getNumericCellValue();
            	System.out.println(empID);
            	
            	String tempDestinationFolder=destinationFolder.replace("firstName-lastName", firstName+lastName);
            
            	File newFile = new File(tempDestinationFolder);
            	
            	FileUtils.copyFile(fileToCopy, newFile);
            	
            	
            	FileInputStream inputStreamTemp = new FileInputStream(new File(tempDestinationFolder));
            	Workbook workbookTemp = new XSSFWorkbook(inputStreamTemp);
            	Sheet tempSheet=workbookTemp.getSheetAt(0);
            	
            	tempSheet.getRow(1).getCell(0).setCellValue((int)empID);
            	/*CellStyle newCellStyle = workbookTemp.createCellStyle();
            	newCellStyle.cloneStyleFrom(nextRow.getCell(0).getCellStyle());
            	newCellStyle.setBorderBottom(nextRow.getCell(0).getCellStyle().getBorderBottom());
            	tempSheet.getRow(1).getCell(0).setCellStyle(newCellStyle);*/
            	
            	
            	tempSheet.getRow(1).getCell(3).setCellValue(email);
            	
            	
            	workbookTemp.getCreationHelper().createFormulaEvaluator().evaluateAll();
            	XSSFFormulaEvaluator.evaluateAllFormulaCells(workbookTemp);
            	
            	FileOutputStream fileOut = new FileOutputStream(tempDestinationFolder);
            	workbookTemp.write(fileOut);
                fileOut.close();
            	
            	workbookTemp.close();
            	inputStreamTemp.close();

            	
            }
            
            
            /*Iterator<Cell> cellIterator = nextRow.cellIterator();
            while (cellIterator.hasNext()) 
            {
                Cell cell = cellIterator.next();
                System.out.print(cell.getCellType());
                switch (cell.getCellType()) {
                    case STRING:
                        System.out.print(cell.getStringCellValue());
                        break;
                    case BOOLEAN:
                        System.out.print(cell.getBooleanCellValue());
                        break;
                    case NUMERIC:
                        System.out.print(cell.getNumericCellValue());
                        break;
                }
                System.out.print(" - ");
            }
            System.out.println();*/
        }
         
        workbook.close();
        inputStream.close();
    }
	
}
