	import net.sf.cglib.asm.CodeAdapter;
	import net.sf.json.JSONException;

	import org.apache.commons.validator.DateValidator;
	import org.apache.poi.hssf.usermodel.HSSFCell;
	import org.apache.poi.hssf.usermodel.HSSFRow;
	import org.apache.poi.hssf.usermodel.HSSFSheet;
	import org.apache.poi.hssf.usermodel.HSSFWorkbook;


	import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
	import org.apache.poi.openxml4j.opc.OPCPackage;
	import org.apache.poi.poifs.filesystem.POIFSFileSystem;
	import org.apache.poi.ss.usermodel.*;


	import org.apache.poi.hssf.usermodel.HSSFCell;
	import org.apache.poi.hssf.usermodel.HSSFRow;
	import org.apache.poi.hssf.usermodel.HSSFSheet;
	import org.apache.poi.hssf.usermodel.HSSFWorkbook;

	import org.apache.poi.xssf.usermodel.XSSFCell;
	import org.apache.poi.xssf.usermodel.XSSFRow;
	import org.apache.poi.xssf.usermodel.XSSFSheet;
	import org.apache.poi.xssf.usermodel.XSSFWorkbook;
	import org.json.JSONArray;
	import org.json.JSONObject;

	import java.io.ByteArrayInputStream;
	import java.io.ByteArrayOutputStream;
	import java.io.File;
	import java.io.FileInputStream;

	import java.text.SimpleDateFormat;
	import java.util.Date;
	import java.sql.*;
	import java.util.List;
	import java.util.ArrayList;
	import com.shm.utils.shmDataStructure;
	import com.sun.mail.iap.ByteArray;

	import javax.ws.rs.Consumes;
	import javax.ws.rs.GET;
	import javax.ws.rs.POST;
	import javax.ws.rs.Path;
	import javax.ws.rs.PathParam;
	import javax.ws.rs.Produces;
	import javax.ws.rs.core.MediaType;

public class ExcelFileRead {



	
		//@MS
		private boolean isRowEmpty(Row row) {
			/*------------------------------------------------------------------------
			/  Desc:
			/	  To check a row in excel file is empty are have data.
			/
			/ Arguments:
			/     Row	row from excel file 
			/	  
			/
			/ Modifications:
			/     Version 1.0(REMS)	29 Jan 2018   Created  by (MS)
			/                     	?? ??? ????   Reviewed by (??)
			/                     	New routine
			/------------------------------------------------------------------------*/
			boolean isRowEmpty = true;
	        for (int c = row.getFirstCellNum(); c < row.getLastCellNum(); c++) {
	            Cell cell = row.getCell(c);
	            if (cell == null || cell.getCellType() == Cell.CELL_TYPE_BLANK)
	            	isRowEmpty =  false;
	            else 
	            	return true;
	        }
	        return isRowEmpty;
	    }
		
		public ArrayList<String> importFile(byte[] excelfile , String UserID){
			/*------------------------------------------------------------------------
			/  Desc:
			/	  Add new record in database from excel file.
			/
			/ Arguments:
			/     arrUserInputsVO	VO containing user's input
			/	  UserID			ID if login user
			/	  strSubmit			Remarks
			/
			/ Modifications:
			/     Version 1.0(REMS)	29 Jan 2018   Created  by (MS)
			/                     	?? ??? ????   Reviewed by (??)
			/                     	New routine
			/------------------------------------------------------------------------*/
			
			
			
			ArrayList<String> faults  = new ArrayList<String>();
			String error = "";
			int successCounter = 0;
			
			ByteArrayInputStream  out = new ByteArrayInputStream(excelfile) ;
			
			try {
			    OPCPackage fs =  OPCPackage.open(out);
			    XSSFWorkbook wb = new XSSFWorkbook(fs);
			    XSSFSheet sheet = wb.getSheetAt(0);
			    XSSFRow row = null;
			    XSSFCell cell;

			    int rows; // No of rows
			    rows = sheet.getPhysicalNumberOfRows();
			    
			    
			    
			    
			    int cols = 0; // No of columns
			    int tmp = 0;
			   // This trick ensures that we get the data properly even if it doesn't start from first few rows
			    for(int i = 0; i < 10 || i < rows; i++) {
			        row = sheet.getRow(i);
			        if(row != null) {
			            tmp = sheet.getRow(i).getPhysicalNumberOfCells();
			            if(tmp > cols) cols = tmp;
			        }
			    }
			    if(cols != 27){
			    	faults.add("file format is incorrect.");
			    	return faults;
			    }
			    
			    for(int r = 1; r < rows; r++) 
			    {
			    	error = "";
			        row = sheet.getRow(r);
			        if(isRowEmpty(row))
			        {
			      
			        for(int c = 0; c < cols; c++) 
			        {
			            	
			            	//cell = row.getCell((short)c);
			                cell = row.getCell(c , org.apache.poi.ss.usermodel.Row.CREATE_NULL_AS_BLANK);
	                           // System.out.println(cell.toString());
			                   
			                    switch(cell.getColumnIndex()) 
				                {
				                    case 0:
				                    	//Gets your first cell data
				                    	cell.toString().trim();
				                    case 1:
				                    	//Gets your second cell data
				                    	cell.toString().trim();
				                    	break;
				                    case 2:
				                    	//Gets your third cell data
				                    	cell.toString().trim();
				                    	break;
				                  
				                }//end switch
			                    
			                    
			            }
			             
			           
			           	
			         
			           // plotList.add(vo);
			      
			        
			    }
			  }
			    
			} catch(Exception ioe) {
			    ioe.printStackTrace();
			    error = ioe.toString();
			}
			
			if(successCounter > 0){
				if(successCounter == 1 ) 
					faults.add( 0,successCounter + " Row Imported Successfully.");
				else
					faults.add( 0, successCounter + " Rows Imported Successfully.");
			}
			if(faults.isEmpty()){
				faults.add("No record Found.");
			}
			return faults;
		}
	



}
