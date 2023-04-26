import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

/*
This program calculates how many email campaigns went towards specific initiatives. 

Time frame: January 1, 2022 – January 1, 2023
*/ 

public class emailCampaignCalculator {
    /**
     * @param args the command line arguments
     * @throws IOException
     * @throws EncryptedDocumentException
     */

     public static void main(String[] args) throws EncryptedDocumentException, IOException {
       
        //Declare variables to be used later to calculate precentages.
        int relief = 0;
        int hr = 0;
        int campus = 0;
        int totalEmails = 0;

        //Access the Excel file.
        Workbook wb = null;

        FileInputStream fins = new FileInputStream(new File("campaigns.xlsx"));
        wb = WorkbookFactory.create(fins);
        Sheet sheet = wb.getSheetAt(0);

        //Iterate through each cell in the Excel file.
        for(Row row : sheet){

            for(Cell cell : row) {
                CellType cellType = cell.getCellType().equals(CellType.FORMULA) ? cell.getCachedFormulaResultType() : cell.getCellType();

                switch (cellType) {

                   case NUMERIC:
                        break;
                        
                    case STRING:
                    
                    //Find the keywords for each email type and add each occurance to the appropriate variable.
                        if (cell.getStringCellValue().contains("T.relief")){
                            relief++;
                        }
            
                        if (cell.getStringCellValue().contains("T.campus")){
                            campus++;
                        }
            
                        if (cell.getStringCellValue().contains("T.HR")){
                            hr++;
                        }
                        
                        break;
                        
                    default:
                        break; 
                       
                }
                
            }        
          
        }

        //Calculate the percentage of emails that each email type makes up.
        totalEmails = 700;
        
        float reliefPercent = (relief * 100) / totalEmails;
        float hrPercent = (hr * 100) / totalEmails;
        float campusPercent = (campus * 100) / totalEmails; 

        //Print the percentage results
        System.out.println("Disaster relief emails: " + reliefPercent + "%.");
        System.out.println("Campus emails: " + campusPercent + "%.");
        System.out.println("HR emails: " + hrPercent + "%.");        
        
    }
}