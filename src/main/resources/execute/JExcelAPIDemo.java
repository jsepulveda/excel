package execute;

/**
 * Main class - Process excel 2003-2007 spreadsheets with several options.
 * 
 * @author juan_sepulveda
 * 
 */

import java.io.File;
import java.io.IOException;

import jxl.Cell;
import jxl.Sheet;
import jxl.Workbook;

import jxl.read.biff.BiffException;

import jxl.write.Label;
import jxl.write.Number;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;

// to read scanf
import java.util.ArrayList;
import java.util.List;
import java.util.Scanner;
//to manage rows with apache poi



public class JExcelAPIDemo
{

	public static void main(String[] args) 
      throws BiffException, IOException, WriteException
   {
	   
	//Reads the files from a directory and prints them out*************************.
//		final String DIRECTORY = "c:/myfiles"; // <-- choose your path
//		final File fileInstance = new File(DIRECTORY); // instantiate class File
//
//		final String[] test = fileInstance.list(); // now list() from File can be used!
//		
//	      
//			   
//			   
//		for (int i = 0, n = test.length; i < n; i++) {
//			
//			System.out.println(test[i]);
//			
//		}
//************************************************************************************
		
		//**********  Prints out the excels to process that are stored in input folder*** 
		final String DIRECTORY = "C:/Users/juan_sepulveda/workspace/ProcessExcel/input";
		final File folderInstance = new File(DIRECTORY); //instantiate class File
		
		final String[] test = folderInstance.list();
		List<String> onlyexcel = new ArrayList<String>();
		
		System.out.println("Excel documents to process (stored in /input folder):");
		
		for (int i = 0, n = test.length; i < n; i++) {
		    	   		
    		if(test[i].contains("xls")){
    			
    			onlyexcel.add(test[i]);
    			
    		}
			
		}
		 for(int i = 0; i < onlyexcel.size(); i++) {
	            System.out.println(onlyexcel.get(i).toString());
	        }
		

	            
	      System.out.println ("************MENU*****************************************************");
	      System.out.println ("1.Introduce a new column in all the spreadsheets.");
	      System.out.println ("2.Introduce a new row in all the spreadsheets.");
	      System.out.println ("3.Remove a tag in all spreadsheets.");
	      System.out.println ("\n");
	      System.out.println ("Please insert a option from the menu:");
	      
	      Scanner scan = new Scanner(System. in);
	      
	      int option = scan.nextInt(); 
	      
		switch (option) {
	      case 1:
	    	  System.out.println ("*********************************************************************");
	          System.out.println ("Option selected: Add column in all the spreadsheets");
	          System.out.println ("*********************************************************************");
				    //Reads excels 2003-2007 and adds a cell on them.**************	  
	         
	          System.out.print ("Please introduce the name of the column to add:");
		      String columnName = scan.next();
		      System.out.print ("Column name to add:"+columnName);
	          
	          
				    for (int i = 0, n = onlyexcel.size(); i < n; i++) { 	  
				//	      I open my spreadsheet located in my workspace
				    	  System.out.println ("\nProcessing excel: "+onlyexcel.get(i).toString());
					      Workbook workbook = Workbook.getWorkbook(new File("input/"+onlyexcel.get(i).toString()));
					      WritableWorkbook copy = Workbook.createWorkbook(new File("output/"+onlyexcel.get(i).toString()),workbook);
					      WritableSheet firstsheet = copy.getSheet(0);
					      int lastColumn = firstsheet.getColumns();
					      System.out.println ("Current number of columns:"+lastColumn);
					      Label column = new Label(lastColumn, 0, columnName); 
					      firstsheet.addCell(column); 
					      copy.write();
					      workbook.close();
					      copy.close();
				    }
		      break;
	      case 2:
	          System.out.println ("2.Introduce a row in the spreadsheet.");
	          
	          
	          
	          for (int i = 0, n = onlyexcel.size(); i < n; i++) { 	  
			      //I open my spreadsheet located in my workspace		
	        	  System.out.println ("\nProcessing excel: "+onlyexcel.get(i).toString());
			      Workbook workbook = Workbook.getWorkbook(new File("input/"+onlyexcel.get(i).toString()));
			      WritableWorkbook copy = Workbook.createWorkbook(new File("output/"+onlyexcel.get(i).toString()), workbook);
			      WritableSheet firstsheet = copy.getSheet(0);
	        	  			      
			      int maxRow = firstsheet.getRows();
			      int maxColumn = firstsheet.getColumns();
			      System.out.println("\n:Last row of the spreadsheet:"+maxRow); 
			      System.out.println("Please introduce the value for the followings columns:");
			      
			      int [] values1 = new int [6];
			      String firstLineFromStdin = "";
			      
			      for (int j = 0; j < values1.length; j++) { 
			    	  
			    	  
			    	  Cell a = firstsheet.getCell(j,0);//gets headers
			    	  System.out.println(":"+a.getContents());
			    	  firstLineFromStdin = new Scanner(System.in).next();
			    	  
			    	  			    	  
			    	//  System.out.println(values1.toString());
			    	  Label label = new Label(j,maxRow,firstLineFromStdin);
				  	    
 				  	 firstsheet.addCell(label);

				  	    

			      }

			      

					      
	                  copy.write();
				      workbook.close();
				      System.out.println ("**EXCEL WRITTEN*****************************");
				      copy.close();
	          }
	          
	          
	          break;
	      case 3:
	    	  System.out.println ("*********************************************************************");
	          System.out.println ("Option selected: Remove a tag in all spreadsheets");
	          System.out.println ("*********************************************************************");
	          System.out.print("Please insert the name of the tag to remove(case sensitive) located at column Tags/tags:\n");
	          String tagToRemove = scan.next();
		      System.out.println ("Tag to remove:\n"+tagToRemove);
		      String myTagbyDefault = "Tags";
		     
		    	
	         
	          
	          for (int i = 0, n = onlyexcel.size(); i < n; i++) { 	  
			      //I open my spreadsheet located in my workspace		
	        	  System.out.println ("\nProcessing excel: "+onlyexcel.get(i).toString());
			      Workbook workbook = Workbook.getWorkbook(new File("input/"+onlyexcel.get(i).toString()));
			      WritableWorkbook copy = Workbook.createWorkbook(new File("output/"+onlyexcel.get(i).toString()), workbook);
			      WritableSheet firstsheet = copy.getSheet(0);
	        	  			      
			      int maxColumn = firstsheet.getColumns();
			      System.out.println("\n"+maxColumn); 
			      
			      Cell cell1 = firstsheet.getCell(1, 5);
	                 //System.out.println(cell1.toString());
	                 
	                 Cell a1 = firstsheet.getCell(0,0);
	                 Cell b2 = firstsheet.getCell(1,1);
	                 Cell c2 = firstsheet.getCell(2,1);
	                 String stringa1 = a1.getContents();
	                 String stringb2 = b2.getContents();
	                 String stringc2 = c2.getContents(); 
	                 System.out.println(stringa1); 
	                 System.out.println(stringb2);
	                 System.out.println(stringc2);
			      
	                 int realposition = findColumnCellByTag(myTagbyDefault,firstsheet);      
	                 System.out.println("HERE:"+realposition);
	    
			    	  
	                 for(int j = 0; j < 1000; j++){
				      	  
			          	  
				      	  Cell a = firstsheet.getCell(realposition,j);
				      	 // System.out.println(a.getContents());
				    	  
									  	  if(a.getContents().equals(tagToRemove)){
									  		  System.out.println ("********FOUND************************************************************");

									  	    Label label = new Label(realposition, j,"");
									  	    //removes cell
									  	    firstsheet.addCell(label);
									  						
									  		  
									  	  }
				
			
			
									  	  	
						}
			    		  

		          
	                  copy.write();
				      workbook.close();
				      copy.close();
	          }
	          
	          break;
	      default: System.out.println ("Close program."); 
	      break;
	      
	      
	      
	  }

	     
	   }

	
	
	public static int findColumnCellByTag(String tag,WritableSheet firstsheet ){
		
	
	
	int position = 0;	
	
				    int maxColumn = 22;
					for(int j = 0; j < maxColumn; j++){
			      	  
			          	  
			      	  Cell a = firstsheet.getCell(j,0);
			      	  System.out.println(a.getContents());
			    	  
								  	  if(a.getContents().equals("Tags")){
								  		  System.out.println ("********FOUND************************************************************");
								  		  position = j;
								  		  Cell tagsCell = firstsheet.getCell(j,0);
								  		  System.out.println(position);
								  		  
								  	  }
			
		
		
								  	  	
					}
	return position;
	}
	
}
