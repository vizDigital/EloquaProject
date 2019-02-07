
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;

import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import connector.eloqua.sdk.Configuration;
import connector.eloqua.sdk.client.LandingPageClient;
import connector.eloqua.sdk.rest.model.Elements;
import connector.eloqua.sdk.rest.model.LandingPage;

public class EloquaMain {
	
	private LandingPageClient lPClient;
	
	public EloquaMain() {
		
		this.lPClient = new LandingPageClient(new Configuration("TechnologyPartnerAssistDigital","Drilon.Kerxhaliu", "123Drilon!@#"));
		
		
		
	     //	Elements<LandingPage> list = this.lPClient.searchForLandingPage("drilon kerxhaliu"); 
		//LandingPage landingPage = list.getElements().get(0);
		//String returnedHtml = landingPage.getHtmlContent().getHtml();
	   // System.out.println("Html Content: " + returnedHtml);
		
	}
		
		public static void main (String[] args) {
			
			EloquaMain eloquaMain = new EloquaMain() ;
			eloquaMain.scanExcel();
			
		}
		
		public void scanExcel() {
			
			String excellFolder = "C:\\Users\\luxottica\\Documents\\Excel\\";
	        
	        try (InputStream inp = new FileInputStream(excellFolder + "EloquaExcel.xlsx")) {

	            Workbook workbook = WorkbookFactory.create(inp);
	            Sheet sheet = workbook.getSheetAt(0);

	            for (int r = 1; r <= sheet.getLastRowNum(); r++) {
	                Row row = sheet.getRow(r);
	                this.processRow(row);
	            }
	            
	            inp.close();
	            
	            FileOutputStream fileOut = new FileOutputStream(excellFolder + "EloquaExcel1.xlsx");
	            workbook.write(fileOut);
	            fileOut.close();

	            // Closing the workbook
	            workbook.close();
	        }
	        catch (IOException ex) {
	            System.out.println("The file could not be read : " + ex.getMessage());
	        }
			
		}

		public void processRow(Row row) {
			
			DataFormatter dataFormatter = new DataFormatter();
			
			String tipologiaCheck = dataFormatter.formatCellValue(row.getCell(1));
			
			   // System.out.println();
		      //  System.out.println("Searching for TipologiaCheck: " + tipologiaCheck);
		       // System.out.println();
		        
		        if(tipologiaCheck.contains("LINK"))
		        	
		        {
		        	
		        	
			
			 
	        String landingPageName = dataFormatter.formatCellValue(row.getCell(0));
	        
	 
	        System.out.println();
	        System.out.println("Searching for page: " + landingPageName);
	        System.out.println();
	       
	        Elements<LandingPage> list = this.lPClient.searchForLandingPage(landingPageName);
	        
	        if(list.getTotal() == 0) {
	            Cell statusCell = row.getCell(3);
	            if(statusCell == null) {
	             statusCell = row.createCell(3);
	         }
	            statusCell.setCellValue("KO - Landing Page non è trovato.");
	            System.out.println("Landing page is not found.");
	            return;
	           }
	            
			
	        // When we have duplicates:
             if(list.getTotal() > 1) {
	         Cell statusCell = row.getCell(3);
	         if(statusCell == null) {
	          statusCell = row.createCell(3);
	      }
	         
	         statusCell.setCellValue("KO - Nome Landing page non univoco.");
	         System.out.println("Landing page is duplicated.");
	         return;
	        } 
	        
	        LandingPage landingPage = list.getElements().get(0);
			String returnedHtml = landingPage.getHtmlContent().getHtml();
			
			System.out.println("Html Content: " + returnedHtml);
			
		    String textCheck = dataFormatter.formatCellValue(row.getCell(4));
		    boolean isFound  = returnedHtml.contains(textCheck);
		    
		    // take value from exccel cell 2
		    
		    String str = dataFormatter.formatCellValue(row.getCell(2));
		   
		   int occorenseCheck = Integer.parseInt(str);
		    
		    System.out.println();
	        System.out.println("Number of Occorrense in EXCEL: " + occorenseCheck);
	        System.out.println();
		    
		    
		  //  int lengthHtml = returnedHtml.length();
		    
		    // matches number of Occorense
	        
		    int numberOccorense = StringUtils.countMatches(returnedHtml,textCheck);
		    
		    
	        
	        System.out.println();
			System.out.println("Number of Occorrense in HTML : " + numberOccorense);
	        System.out.println();
		    
	      //  boolean isOccorenseEqual = occorenseCheck.equals(numberOccorense) ;
	        
	        if(occorenseCheck == numberOccorense)
	        	
	        {
	        	System.out.println("Are Equals");
	        	
	        	Cell statusCell = row.getCell(2);
	            if(statusCell == null) {
	             statusCell = row.createCell(2);
	            }
	             statusCell.setCellValue(" OK ! Html:NO => "+numberOccorense+" Excel:NO => " +occorenseCheck);
	             System.out.println("Are Equals");
		           
	             return;
	        }else {
	        	Cell statusCell = row.getCell(2);
	            if(statusCell == null) {
	             statusCell = row.createCell(2);
	            }
	        	
	        	statusCell.setCellValue("Attenzione! Html:NO => "+numberOccorense+" Excel:NO => " +occorenseCheck);
	        	
	        	System.out.println("Attenzione !");
	            
	        	 
	        }
	        
	        
	        /*  
		    for (int i=0;i<lengthHtml;i++) {
		    
		    boolean  numberOccorense = returnedHtml.contains(textCheck);
		    
		    if(numberOccorense == false) {
		    
		    System.out.println("Boolean "+numberOccorense);
		    
		    }
		    
		    }
		    
		   
		    /*    int count = 0;
		    
		   
		     for (int i=0;i<lengthHtml;i++) {
		     if(returnedHtml.contains(textCheck)) {
		    	 
		    	 count ++;
		     }
		     
		     }
		     
		     System.out.println("The number Link found: "+count); */

		     Cell statusCell = row.getCell(3);
		     if(statusCell == null) {
		      statusCell = row.createCell(3);
		 }
		    
		    	 
		     if(isFound == true) {
		      System.out.println("The text was found.");
		      statusCell.setCellValue("OK");
		      
		      
		      
		      
		      
		        } 
		     
		    else {
		      System.out.println("The text was not found.");
		     statusCell.setCellValue("KO");
		    
		     }
		     
		     
			if (list.getTotal()>0) {
				 
			System.out.println("Html Content: " + list.getTotal());
				
			}
				
				 
			
		}
		        
		        else {
		        	System.out.println("Attention!");
		        }
		        
		}
		}
