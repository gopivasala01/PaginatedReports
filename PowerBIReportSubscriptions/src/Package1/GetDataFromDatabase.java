package Package1;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.ResultSet;
import java.sql.Statement;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class GetDataFromDatabase 
  {
	static File file;
	static FileInputStream fis;
	static FileOutputStream fos;
	static XSSFWorkbook workbook;
	static XSSFSheet sheet;
	
     public static void getData() throws IOException 
     {
    	// FileOutputStream outputStream =null;
    	// try
    	// {
    	file  = new File("C:\\Gopi\\Projects\\Property ware\\Paginated Reports\\Agents Info.xlsx");
    	//if(file.exists())
    	//{
    		
    	//fis = new FileInputStream(file);
    	//workbook = new XSSFWorkbook(fis);
    	//}
    	//else 
    	//{
    		//file.delete();
    		//file  = new File("C:\\Gopi\\Projects\\Property ware\\Paginated Reports\\Agents Info.xlsx");
    	//	 outputStream = new FileOutputStream(new File("C:\\Gopi\\Projects\\Property ware\\Paginated Reports\\Agents Info.xlsx"));
    	//	outputStream.close();
    	//}
    	// }
    	// catch(Exception e)
    	// {
    		 //file.createNewFile();
    		 //file  = new File("C:\\Gopi\\Projects\\Property ware\\Paginated Reports\\Agents Info.xlsx");
    	 //}
    	 fis = new FileInputStream(file);
    	 workbook = new XSSFWorkbook(fis);
    	sheet = workbook.getSheetAt(0);
    	if(sheet !=null)
    	{
    		workbook.removeSheetAt(0);
    		sheet = workbook.createSheet();
    	}
	        String name = "";
	        String email ="";
	        String Company ="";
	        String RLM ="";
	        String RLMEmail ="";
	        int rowNumber =0;
	        try {
	            String connectionUrl = "jdbc:sqlserver://azrsrv001.database.windows.net;databaseName=HomeRiverDB;user=service_sql02;password=xzqcoK7T";
	            Connection con = null;
	            Statement stmt = null;
	            ResultSet rs = null;
	                //Class.forName("com.microsoft.sqlserver.jdbc.SQLServerDriver");
	                con = DriverManager.getConnection(connectionUrl);
	                String SQL = "WITH CTE As\r\n"
	                		+ "(\r\n"
	                		+ "Select MAX(ID)\r\n"
	                		+ " as ID,  PrimaryLeasingAgent from VW_Units where [Listing Agent (Name/Phone/Email)] is not NULL and PrimaryLeasingAgent is not NULL\r\n"
	                		+ "and PrimaryLeasingAgent <>'' and [Listing Agent (Name/Phone/Email)]<>'' and [Listing Agent (Name/Phone/Email)] like '%@%' group by PrimaryLeasingAgent\r\n"
	                		+ "),\r\n"
	                		+ "CTE2 as\r\n"
	                		+ "(\r\n"
	                		+ "Select C.PrimaryLeasingAgent, [Listing Agent (Name/Phone/Email)],Company from VW_Units U inner join CTE C on U.ID = C.ID --order by C.PrimaryLeasingAgent\r\n"
	                		+ ")\r\n"
	                		+ "Select C.*,R.RLM, R.RLMEmail from CTE2 C inner join RegionLeaders R on C.Company = R.Region order by PrimaryLeasingAgent";
	                stmt = con.createStatement();
	                rs = stmt.executeQuery(SQL);
	                Row row = sheet.createRow(0);
	                Cell cell = row.createCell(0);
	                cell.setCellValue("Agent Name");
	                
	               // Row row2 = sheet.createRow(0);
	                Cell cell2 = row.createCell(1);
	                cell2.setCellValue("Raw Email");
	                
	               // Row row2 = sheet.createRow(0);
	                Cell cell3 = row.createCell(2);
	                cell3.setCellValue("Status");
	                while (rs.next())
	                {
	                	name =  (String) rs.getObject(1);
	                	email = (String) rs.getObject(2);
	                	Company = (String) rs.getObject(3);
	                	RLM = (String) rs.getObject(4);
	                	RLMEmail = (String) rs.getObject(5);
	                	
	                	Row row2 = sheet.createRow(rowNumber++);
		                Cell cell4 = row2.createCell(0);
		                cell4.setCellValue(name);
		                
		                Cell cell5 = row2.createCell(1);
		                cell5.setCellValue(email);
		                
		                Cell cell6 = row2.createCell(2);
		                cell6.setCellValue(Company);
		                
		                Cell cell8 = row2.createCell(3);
		                cell8.setCellValue(RLM);
		                
		                Cell cell9 = row2.createCell(4);
		                cell9.setCellValue(RLMEmail);
		                
		                Cell cell7 = row2.createCell(5);
		                cell7.setCellValue("Pending");
	                    System.out.println(name);
	                    System.out.println(email);
	                    System.out.println(RLM);
	                }
	                rs.close();
	                
	                fos = new FileOutputStream(file);
	                workbook.write(fos);
	                workbook.close();
	                fos.close();
	        }
	        catch (Exception e) 
	        {
              e.printStackTrace();
	           // return "ERROR while retrieving data: " + e.getMessage();
	        }
	       // return output;
	    
   }
   
   
   
  }


