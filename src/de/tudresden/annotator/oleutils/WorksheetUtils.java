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
		
		if(worksheetAutomation==null)
			return false;
		
		int[] visiblePropertyIds = worksheetAutomation.getIDsOfNames(new String[]{"Visible"});			
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
			System.out.println("\"Activate\" method not found for \"Worksheet\" object!");
			return false;
		}		

		Variant result = worksheetAutomation.invoke(activateMethodsIds[0]);
		if(result==null)
			return false;
		
		result.dispose();
		return true;
	}
	
	
	/**
	 * Make all data in the worksheet visible. Also, invoking this method will remove filters on the sheet data.  
	 * @param worksheetAutomation an OleAutomation for accessing the Worksheet OLE object
	 * @return true if operation was successful, false otherwise
	 */
	public static boolean showAllWorksheetData(OleAutomation worksheetAutomation){
		
		int[] showAllDataMethodsIds = worksheetAutomation.getIDsOfNames(new String[]{"ShowAllData"});	
		if (showAllDataMethodsIds == null){			
			System.out.println("\"ShowAllData\" method not found for \"Worksheet\" object!");
			return false;
		}		

		Variant result = worksheetAutomation.invoke(showAllDataMethodsIds[0]);
		if(result==null)
			return false;
		
		result.dispose();
		return true;
	}
	
	
	/**
	 * Get the specified range automation. The address of the top_left_cell and down_right_cell have to be provided.
	 * For single cell ranges the address of the down_right_cell is NULL or the same as the top_left_cell    
	 * @param worksheetAutomation an OleAutomation object for accessing the Worksheet OLE object
	 * @param topLeftCell address of top left cell (e.g., "A1" or "$A$1" )
	 * @param downRightCell address of down right cell (e.g., "C3" or "$C$3" )
	 * @return an OleAutomation that provides access the specified range 
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
		
		if(rangeVariant==null || rangeVariant.getType()==0){
			return null;
		}
			
		OleAutomation rangeAutomation = rangeVariant.getAutomation();
		for (Variant arg : args) {
			arg.dispose();
		}
		rangeVariant.dispose();
		
		return rangeAutomation;
	}
	
	
	/**
	 * Get the specified range automation given its address. 
	 * This method supports simple ranges having only one area (i.e., not multi-selection)
	 * @param worksheetAutomation an OleAutomation object for accessing the Worksheet OLE object
	 * @param rangeAddress a string representing the address the range (e.g., "$A$1:$C$2", "$D$4", "B2:G6")
	 * @return an OleAutomation that provides access the specified range 
	 */
	public static OleAutomation getRangeAutomation(OleAutomation worksheetAutomation, String rangeAddress){
		
		String[] subStrings = rangeAddress.split(":");
		String topLeftCell = subStrings[0];
		String downRightCell = null;
		if (subStrings.length == 2)
			downRightCell = subStrings[1];
		
		return getRangeAutomation(worksheetAutomation, topLeftCell, downRightCell);
	}
	
	
	/**
	 * Get the OleAutomation for the specified multi-selection range automation 
	 * @param worksheetAutomation an OleAutomation object for accessing the Worksheet OLE object
	 * @param multiSelectionRange a string that represents the address of the multi-selection range (Ex. "$A$1:$C$2, $A$4:$C$4, $D$6" ) 
	 * @return an OleAutomation that provides access to the multi-selection range 
	 */
	public static OleAutomation getMultiSelectionRangeAutomation(OleAutomation worksheetAutomation, String multiSelectionRange){
		
		// get the OleAutomation object for the multi-selection (multi-area) range 
		int[] rangePropertyIds = worksheetAutomation.getIDsOfNames(new String[]{"Range"});
		
		Variant[] args = new Variant[1];
		args[0] = new Variant(multiSelectionRange);

		Variant rangeVariant = worksheetAutomation.getProperty(rangePropertyIds[0], args);
		OleAutomation rangeAutomation = rangeVariant.getAutomation();
		
		for (Variant arg : args) {
			arg.dispose();
		}
		rangeVariant.dispose();
		
		return rangeAutomation;
	}
	

	/**
	 * Get the used range for the given worksheet
	 * @param worksheetAutomation an OleAutomation object for accessing the Worksheet OLE object
	 * @return an OleAutomation for accessing the used range
	 */
	public static OleAutomation getUsedRange(OleAutomation worksheetAutomation){
		
		int[] usedRangePropertyIds = worksheetAutomation.getIDsOfNames(new String[]{"UsedRange"});	
		Variant usedRangeVariant = worksheetAutomation.getProperty(usedRangePropertyIds[0]);
		OleAutomation usedRangeAutomation = usedRangeVariant.getAutomation();
		usedRangeVariant.dispose();
		
		return usedRangeAutomation;
	}
	
	
	/**
	 * Get a cell from the specified worksheet given the row and the column number. 
	 * @param worksheetAutomation an OleAutomation to access the worksheet that contains the cell
	 * @param row an integer that represents the row number (index)
	 * @param column an integer that represents the column number (index)
	 * @return an OleAutomation that provides access to the cell. 
	 */
	public static OleAutomation getCell(OleAutomation worksheetAutomation, int row, int column){
		
		int[] cellsPropertyIds = worksheetAutomation.getIDsOfNames(new String[]{"Cells"}); 
		
		Variant[] args = new Variant[2];
		args[0] = new Variant(row);
		args[1] = new Variant(column);
		
		Variant cellsVariant = worksheetAutomation.getProperty(cellsPropertyIds[0], args);	
		OleAutomation cellAutomation = cellsVariant.getAutomation();
		cellsVariant.dispose();
		
		args[0].dispose();
		args[1].dispose();
		
		return cellAutomation; 
	}
	
	
	/**
	 * Get collection of columns in the worksheet 
	 * @param worksheetAutomation an OleAutomation to access the worksheet that contains the cell
	 * @return an OleAutomation that provides access to the collection of columns in the worksheet 
	 */
	public static OleAutomation getRangeColumns(OleAutomation worksheetAutomation){
				
		int[] columnsPropertyIds = worksheetAutomation.getIDsOfNames(new String[]{"Columns"}); 
		Variant columnsPropertyVariant = worksheetAutomation.getProperty(columnsPropertyIds[0]);	
		
		OleAutomation columnsAutomation =  columnsPropertyVariant.getAutomation();
		columnsPropertyVariant.dispose();
		
		return columnsAutomation;
	}
	

	/**
	 * Get a specific column in the worksheet 
	 * @param worksheetAutomation an OleAutomation to access the worksheet that contains the cell
	 * @return an OleAutomation that provides access to the range that represents the entire specified column in the worksheet
	 */
	public static OleAutomation getRangeColumn(OleAutomation worksheetAutomation, String column){
				
		int[] columnsPropertyIds = worksheetAutomation.getIDsOfNames(new String[]{"Columns"}); 
		
		Variant[] args= new Variant[1];
		args[0] = new Variant(column);
		
		Variant columnPropertyVariant = worksheetAutomation.getProperty(columnsPropertyIds[0], args);	
		args[0].dispose();
		
		OleAutomation columnAutomation =  columnPropertyVariant.getAutomation();
		columnPropertyVariant.dispose();
		
		return columnAutomation;
	}
	
	
	/**
	 * Get collection of rows in the worksheet 
	 * @param worksheetAutomation an OleAutomation to access the worksheet that contains the cell
	 * @return an OleAutomation that provides access to the collection of rows in the worksheet 
	 */
	public static OleAutomation getRangeRows(OleAutomation worksheetAutomation){
		
		int[] rowsPropertyIds = worksheetAutomation.getIDsOfNames(new String[]{"Rows"}); 
		Variant rowsPropertyVariant = worksheetAutomation.getProperty(rowsPropertyIds[0]);	
		OleAutomation rowsAutomation = rowsPropertyVariant.getAutomation();
		rowsPropertyVariant.dispose();
		
		return rowsAutomation;
	}
	
	
	/**
	 * Get a specific row from the worksheet
	 * @param worksheetAutomation an OleAutomation to access the worksheet that contains the cell
	 * @return an OleAutomation that provides access to the range that represents the entire specified row in the worksheet
	 */
	public static OleAutomation getRangeRow(OleAutomation worksheetAutomation, int row){
		
		int[] rowsPropertyIds = worksheetAutomation.getIDsOfNames(new String[]{"Rows"}); 

		Variant[] args= new Variant[1];
		args[0] = new Variant(row);
		
		Variant rowPropertyVariant = worksheetAutomation.getProperty(rowsPropertyIds[0], args);
		args[0].dispose();
		
		OleAutomation rowAutomation = rowPropertyVariant.getAutomation();
		rowPropertyVariant.dispose();
		
		return rowAutomation;
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
		// you can find the documentation of this OLE method at 
		// https://msdn.microsoft.com/EN-US/library/ff840611.aspx
		int[] protectMethodIds = worksheetAutomation.getIDsOfNames(new String[]{"Protect", 
				"AllowFormattingColumns", "AllowFormattingRows"});
		
		if (protectMethodIds == null) {
			System.out.println("Method \"Protect\" of \"Worksheet\" OLE Object is not found!");
			return false;
		}else{
			Variant[] args = new Variant[2];
			args[0] = new Variant(true); // allow user to resize columns
			args[1] = new Variant(true); // allow user to resize rows
			
			int argsIds[] = Arrays.copyOfRange(protectMethodIds, 1, protectMethodIds.length);
			
			Variant result = worksheetAutomation.invoke(protectMethodIds[0], args, argsIds);	
			if(result==null){
				System.err.println("The worksheet.protect method returned null");
				System.err.println(WorksheetUtils.getWorksheetName(worksheetAutomation));
				System.exit(1);
				// return false;
			}
			
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
			System.err.println("The worksheet.unprotect method returned null");
			System.err.println(WorksheetUtils.getWorksheetName(worksheetAutomation));
			System.exit(1);
			//return false;
		}	
		
		result.dispose();
		return true;
	}
	
	/**
	 * Export the data in the given worksheet as a CSV file 
	 * @param worksheetAutomation an OleAutomation for accessing the Worksheet OLE object
	 * @param filePath the path of the file to save the data 
	 * @return true if the operation was successful, false otherwise
	 */
	public static boolean saveAsCSV(OleAutomation worksheetAutomation, String filePath){
		
		int[] saveAsMethodIds = worksheetAutomation.getIDsOfNames(new String[]{"SaveAs", "FileName", "FileFormat"});	
		
		Variant[] args = new Variant[2];
		args[0] = new Variant(filePath);
		args[1] = new Variant(6); // xlCSV = 6 
		
		int argsIds[] = Arrays.copyOfRange(saveAsMethodIds, 1, saveAsMethodIds.length); 
		Variant result = worksheetAutomation.invoke(saveAsMethodIds[0], args, argsIds);
		
		if(result==null){
			System.err.println("The Worksheet.SaveAs method returned null");
			System.err.println(WorksheetUtils.getWorksheetName(worksheetAutomation));
			System.exit(1);
			// return false;
		}
		
		for (Variant arg : args) {
			arg.dispose();
		}
		result.dispose();
		
		return true;
	}
	
//	/**
//	 * Delete the specified worksheet
//	 * @param worksheetAutomation an OleAutomation for accessing the Worksheet OLE object
//	 * @return true if the worksheet was successfully deleted, false otherwise
//	 */
//	public static boolean deleteWorksheet(OleAutomation worksheetAutomation){
//		
//		int[] deleteMethodIds = worksheetAutomation.getIDsOfNames(new String[]{"Delete"});
//		Variant result = worksheetAutomation.invoke(deleteMethodIds[0]);
//		
//		if(result!=null){		
//			boolean isSuccess = result.getBoolean();
//			result.dispose();
//
//			return isSuccess;
//		}else{
//			return false;
//		}
//	}

}
