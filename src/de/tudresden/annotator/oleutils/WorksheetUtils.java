/**
 * 
 */
package de.tudresden.annotator.oleutils;

import java.util.Arrays;

import org.eclipse.swt.ole.win32.OleAutomation;
import org.eclipse.swt.ole.win32.Variant;

/**
 * @author Elvis Koci
 */
public class WorksheetUtils {
		
	/**
	 * Get the name of the given worksheet
	 * @param worksheetAutomation an OleAutomation for accessing the Worksheet OLE object
	 * @return the name of the worksheet
	 */
	public static String getWorksheetName(OleAutomation worksheetAutomation){
		
		int[] namePropertyIds = worksheetAutomation.getIDsOfNames(new String[]{"Name"});	
		if (namePropertyIds == null){			
			System.out.println("\"Name\" property not found for \"Worksheet\" object!");
			return null;
		}		
		
		Variant nameVariant = worksheetAutomation.getProperty(namePropertyIds[0]);
		if (nameVariant == null) {
			System.out.println("\"Name\" variant is null!");
			return null;
		}
		
		String worksheetName = nameVariant.getString();
		nameVariant.dispose();
		
		return worksheetName;
	}
	
	
	/**
	 * Get the index of the given worksheet
	 * @param worksheetAutomation an OleAutomation for accessing the Worksheet OLE object
	 * @return the index number of the worksheet in the collection Workbook.Worksheets 
	 */
	public static long getWorksheetIndex(OleAutomation worksheetAutomation){
		
		int[] indexPropertyIds = worksheetAutomation.getIDsOfNames(new String[]{"Index"});	
		if (indexPropertyIds == null){			
			System.out.println("\"Index\" property not found for \"Worksheet\" object!");
			return 0;
		}		
		
		Variant indexVariant = worksheetAutomation.getProperty(indexPropertyIds[0]);
		if (indexVariant == null) {
			System.out.println("\"Index\" variant is null!");
			return 0;
		}
		
		long worksheetIndex = indexVariant.getLong();
		indexVariant.dispose();
		
		return worksheetIndex;
	}
	
	/**
	 * Get the worksheet automation from the embedded workbook based on the given name  
	 * @param workbookAutomation an OleAutomation that provides access to the functionalities  of the Worksheets OLE object.
	 * @param sheetName the name of the worksheet
	 * @return
	 */
	public static OleAutomation getWorksheetAutomationByName(OleAutomation worksheetsAutomation, String sheetName){
		OleAutomation worksheetAutomation = CollectionsUtils.getItemByName(worksheetsAutomation, sheetName);
		return worksheetAutomation;
	}	
	
	
	/**
	 * Get the worksheet automation from the embedded workbook based on the index number  
	 * @param workbookAutomation an OleAutomation that provides access to the functionalities of the Worksheets OLE object.
	 * @param index an integer that represents the index of the worksheet in the collection of worksheets.
	 * @return an OleAutomation to access worksheet object functionalities
	 */
	public static OleAutomation getWorksheetAutomationByIndex(OleAutomation worksheetsAutomation, int index){	
		OleAutomation worksheetAutomation = CollectionsUtils.getItemByIndex(worksheetsAutomation, index);
		return worksheetAutomation;
	}	
	
	
	/**
	 * Get the specified range automation. The address of the top left cell and down right cell have to be provided.
	 * @param worksheetAutomation an OleAutomation object for accessing the Worksheet OLE object
	 * @param topLeftCell address of top left cell (e.g., "A1" or "$A$1" )
	 * @param downRightCell address of down right cell (e.g., "C3" or "$C$3" )
	 * @return
	 */
	public static OleAutomation getRangeAutomation(OleAutomation worksheetAutomation, String topLeftCell, String downRightCell){
		
		// get the OleAutomation object for the selected range 
		int[] rangePropertyIds = worksheetAutomation.getIDsOfNames(new String[]{"Range"});
		
		Variant[] args;
		if(downRightCell!=null && downRightCell.length()>1){
			args = new Variant[2];
			args[0] = new Variant(topLeftCell);
			args[1] = new Variant(downRightCell);
		}else{
			args = new Variant[1];
			args[0] = new Variant(topLeftCell);
		}
		
		Variant rangeVariant = worksheetAutomation.getProperty(rangePropertyIds[0],args);
		OleAutomation rangeAutomation = rangeVariant.getAutomation();
		for (Variant arg : args) {
			arg.dispose();
		}
		rangeVariant.dispose();
		
		return rangeAutomation;
	}
	
	/**
	 * Get the OleAutomation object for the "Shapes" property of the given worksheet  
	 * @param worksheetAutomation an OleAutomation for accessing the Worksheet OLE object
	 * @return an OleAutomation to access the Shapes of the worksheet 
	 */
	public static OleAutomation getWorksheetShapes(OleAutomation worksheetAutomation){
		
		int[] shapesPropertyIds = worksheetAutomation.getIDsOfNames(new String[]{"Shapes"});	
		if (shapesPropertyIds == null){			
			System.out.println("\"Shapes\" property not found for \"Worksheet\" object!");
			return null;
		}		
		
		Variant shapesVariant = worksheetAutomation.getProperty(shapesPropertyIds[0]);
		if (shapesVariant == null) {
			System.out.println("\"Shapes\" variant is null!");
			return null;
		}
		
		OleAutomation worksheetShapes = shapesVariant.getAutomation();
		shapesVariant.dispose();
		
		return worksheetShapes;		
	}
	
	
	/**
	 * Protect all worksheets that are part of the embedded workbook
	 * @param worksheetsAutomation an OleAutomation for accessing the Worksheets OLE object (Represents a collection of worksheets in a workbook)
	 * @return true if operation succeeded, false otherwise
	 */
	public static boolean protectWorksheets(OleAutomation worksheetsAutomation){
		
		int count = CollectionsUtils.getNumberOfObjectsInOleCollection(worksheetsAutomation);
		
		int i; 
		boolean isSuccess=true; 
		for (i = 1; i <= count; i++) {
		
			OleAutomation nextWorksheetAutomation = CollectionsUtils.getItemByIndex(worksheetsAutomation, i);					
			if(!protectWorksheet(nextWorksheetAutomation)){
				System.out.println("ERROR: Could not protect one of the workbooks!");
				isSuccess=false;			
			}	
			nextWorksheetAutomation.dispose();	
			if(!isSuccess){
				break;
			}
		}	
		
		if(!isSuccess){
			for(int j=1; j<i;j++){
				OleAutomation nextWorksheetAutomation =  CollectionsUtils.getItemByIndex(worksheetsAutomation, j);
				unprotectWorksheet(nextWorksheetAutomation);
				nextWorksheetAutomation.dispose();
			}
			worksheetsAutomation.dispose();
			return false;
		}
		
		worksheetsAutomation.dispose();
		return true;
	}
	
	/**
	 * Protect the data, formating, and structure of the specified worksheet
	 * @param worksheetAutomation an OleAutomation for accessing the Worksheet OLE object
	 * @return true if operation succeeded, false otherwise
	 */
	public static boolean protectWorksheet(OleAutomation worksheetAutomation){
		
		// get the id of the "Protect" method and the considered parameters
		// you can find the documentation of this OLE method here https://msdn.microsoft.com/EN-US/library/ff840611.aspx
		int[] protectMethodIds = worksheetAutomation.getIDsOfNames(new String[]{"Protect", "DrawingObjects", "Contents",  
					"Scenarios", "AllowFormattingCells", "AllowFormattingColumns", "AllowFormattingRows", 
					"AllowInsertingColumns", "AllowInsertingRows","AllowInsertingHyperlinks", "AllowDeletingColumns",
					"AllowDeletingRows", "AllowSorting", "AllowFiltering", "AllowUsingPivotTables" });
		
		if (protectMethodIds == null) {
			System.out.println("Method \"Protect\" of \"Worksheet\" OLE Object is not found!");
			return false;
		}else{
			Variant[] args = new Variant[14];
			args[0] = new Variant(true);
			args[1] = new Variant(true);
			args[2] = new Variant(true);
			args[3] = new Variant(false);
			args[4] = new Variant(true); // allow user to resize columns
			args[5] = new Variant(true); // allow user to resize rows
			args[6] = new Variant(false);
			args[7] = new Variant(false);
			args[8] = new Variant(false);
			args[9] = new Variant(false);
			args[10] = new Variant(false);
			args[11] = new Variant(false);
			args[12] = new Variant(false);
			args[13] = new Variant(false);
			
			Variant result = worksheetAutomation.invoke(protectMethodIds[0],args,Arrays.copyOfRange(protectMethodIds, 1, protectMethodIds.length));	
			result.dispose();
			for (Variant arg: args) {
				arg.dispose();
			}
		}				
		worksheetAutomation.dispose();	
		return true;
	}
	
	
	/**
	 * Unprotect the specified worksheet 
	 * @param worksheetAutomation an OleAutomation for accessing the Worksheet OLE object
	 * @return true if operation succeeded, false otherwise
	 */
	public static boolean unprotectWorksheet(OleAutomation worksheetAutomation){
		
		// get the id of the "Unprotect" method for worksheet OLE object 
		int[] unprotectMethodIds = worksheetAutomation.getIDsOfNames(new String[]{"Unprotect"});
		if(unprotectMethodIds==null){
			System.out.println("Method \"Unprotect\" of \"Worksheet\" OLE Object is not found!");
			return false;
		}
		
		// invoke the unprotect method  
		Variant result = worksheetAutomation.invoke(unprotectMethodIds[0]);

		if(result==null){
			return false;
		}	
		
		result.dispose();
		return true;
	}
}
