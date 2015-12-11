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
	 * set the name of the given worksheet
	 * @param worksheetAutomation an OleAutomation for accessing the Worksheet OLE object
	 * @param a string that represents the name to set for the worksheet
	 * @return true if operation was successful, false otherwise
	 */
	public static boolean setWorksheetName(OleAutomation worksheetAutomation, String name){
		
		int[] namePropertyIds = worksheetAutomation.getIDsOfNames(new String[]{"Name"});	
		if (namePropertyIds == null){			
			System.out.println("\"Name\" property not found for \"Worksheet\" object!");
			return false;
		}		
		
		Variant nameVariant = new Variant(name); 
		boolean isSuccess = worksheetAutomation.setProperty(namePropertyIds[0], nameVariant);
		nameVariant.dispose();
		
		return isSuccess;
	}
	
	
	/**
	 * Get the index of the given worksheet
	 * @param worksheetAutomation an OleAutomation for accessing the Worksheet OLE object
	 * @return the index number of the worksheet in the collection Workbook.Worksheets 
	 */
	public static int getWorksheetIndex(OleAutomation worksheetAutomation){
		
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
		
		int worksheetIndex = indexVariant.getInt();
		indexVariant.dispose();
		
		return worksheetIndex;
	}
	
	
	/**
	 * set worksheet visibility
	 * @param worksheetAutomation an OleAutomation for accessing the Worksheet OLE object
	 * @param true to set worksheet visible, false to hide it
	 * @return true if operation was successful, false otherwise
	 */
	public static boolean setWorksheetVisibility(OleAutomation worksheetAutomation, boolean visible){
		
		int[] visiblePropertyIds = worksheetAutomation.getIDsOfNames(new String[]{"Visible"});	
		if (visiblePropertyIds == null){			
			System.out.println("\"Visible\" property not found for \"Worksheet\" object!");
			return false;
		}		
		
		Variant visiblePropertyVariant = new Variant(visible); 
		boolean isSuccess = worksheetAutomation.setProperty(visiblePropertyIds[0], visiblePropertyVariant);
		visiblePropertyVariant.dispose();
		
		return isSuccess;
	}
	
	
	/**
	 * Make the given worksheet the active sheet.
	 * @param worksheetAutomation an OleAutomation for accessing the Worksheet OLE object
	 * @return true if operation was successful, false otherwise
	 */
	public static boolean makeWorksheetActive(OleAutomation worksheetAutomation){
		
		int[] activateMethodsIds = worksheetAutomation.getIDsOfNames(new String[]{"Activate"});	
		if (activateMethodsIds == null){			
			System.out.println("\"Activate\" property not found for \"Worksheet\" object!");
			return false;
		}		

		Variant result = worksheetAutomation.invoke(activateMethodsIds[0]);
		if(result==null)
			return false;
		
		result.dispose();
		return true;
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
	
	public static OleAutomation getUsedRange(OleAutomation worksheetAutomation){
		
		int[] usedRangePropertyIds = worksheetAutomation.getIDsOfNames(new String[]{"UsedRange"});
		
		Variant usedRangeVariant = worksheetAutomation.getProperty(usedRangePropertyIds[0]);
		OleAutomation usedRangeAutomation = usedRangeVariant.getAutomation();
		usedRangeVariant.dispose();
		
		return usedRangeAutomation;
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
	
	
	public static boolean saveAsCSV(OleAutomation worksheetAutomation, String fileName){
		
		int[] saveAsMethodIds = worksheetAutomation.getIDsOfNames(new String[]{"SaveAs", "FileName", "FileFormat"});	
		
		Variant[] args = new Variant[2];
		args[0] = new Variant(fileName);
		args[1] = new Variant(6); // xlCSV = 6 
		
		int argsIds[] = Arrays.copyOfRange(saveAsMethodIds, 1, saveAsMethodIds.length); 
		Variant result = worksheetAutomation.invoke(saveAsMethodIds[0], args, argsIds);
		
		if(result==null)
			return false;
					
		for (Variant arg : args) {
			arg.dispose();
		}
		result.dispose();
		
		return true;
	}
}
