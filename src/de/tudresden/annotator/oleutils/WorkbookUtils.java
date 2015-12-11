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
public class WorkbookUtils {
	
	/**
	 * Get the name of the given workbook
	 * @param workbookAutomation an OleAutomation that provides access to the functionalities of the Workbook OLE object 
	 * @return the name of the workbook
	 */
	public static String getWorkbookName(OleAutomation workbookAutomation){
		
		int[] namePropertyIds = workbookAutomation.getIDsOfNames(new String[]{"Name"});	
		if (namePropertyIds == null){			
			System.out.println("\"Name\" property not found for \"Workbook\" object!");
			return null;
		}		
		
		Variant nameVariant = workbookAutomation.getProperty(namePropertyIds[0]);
		if (nameVariant == null) {
			System.out.println("\"Name\" variant is null!");
			return null;
		}
		
		String workbookName = nameVariant.getString();
		nameVariant.dispose();
		
		return workbookName;
	}
	

	/**
	 * Get the Worksheets automation
	 * @param workbookAutomation an OleAutomation that provides access to the functionalities of the Workbook OLE object 
	 * @return an OleAutomation to access the Worksheets collection
	 */
	public static OleAutomation getWorksheetsAutomation(OleAutomation workbookAutomation){
		
		// get ID of Worksheets property
		int[] worksheetsObjectIds = workbookAutomation.getIDsOfNames(new String[]{"Worksheets"});
		if (worksheetsObjectIds == null) {
			System.out.println("Property \"Worksheets\" was not found for the given Workbook OLE object!");
			return null;
		}
		
		// get property using the ID 
		Variant worksheetsVariant =  workbookAutomation.getProperty(worksheetsObjectIds[0]);	
		if(worksheetsVariant == null){
			System.out.println("\"Worksheets\" variant is null!");
			return null;		
		}
		// get automation from the Worksheets variant
		OleAutomation worksheetsAutomation = worksheetsVariant.getAutomation();
		worksheetsVariant.dispose();
		
		return worksheetsAutomation;
	}
	
	
	/**
	 * Get the active worksheet automation using the "ActiveSheet" property. 
	 * @param workbookAutomation an OleAutomation that provides access to the functionalities of the Workbook OLE object
	 * @return an OleAutomation for the ActiveWorksheet
	 */
	public static OleAutomation getActiveWorksheetAutomation(OleAutomation workbookAutomation){
		
		int[] worksheetIds = workbookAutomation.getIDsOfNames(new String[]{"ActiveSheet"});	
		if (worksheetIds == null){			
			System.out.println("\"ActiveSheet\" property not found for the given OleAutomation object!");
			return null;
		}		
		Variant worksheetVariant = workbookAutomation.getProperty(worksheetIds[0]);
		if (worksheetVariant == null) {
			System.out.println("Workbook variant is null!");
			return null;
		}		
		OleAutomation worksheetAutomation = worksheetVariant.getAutomation();
		worksheetVariant.dispose();
		
		return worksheetAutomation;
	}
	
	
	/**
	 * Get the worksheet automation from the embedded workbook based on the given name  
	 * @param workbookAutomation an OleAutomation that provides access to the functionalities of a Workbook OLE object
	 * @param sheetName the name of the worksheet
	 * @return
	 */
	public static OleAutomation getWorksheetAutomationByName(OleAutomation workbookAutomation, String sheetName){
		
		OleAutomation worksheetsAutomation = getWorksheetsAutomation(workbookAutomation);
		if(worksheetsAutomation==null){
			System.out.println("ERROR: Could not retrieve Worksheets automation!!!");
			System.exit(1);
		}	
		
		OleAutomation worksheetAutomation = CollectionsUtils.getItemByName(worksheetsAutomation, sheetName, false);
		worksheetsAutomation.dispose();
		
		return worksheetAutomation;
	}	
		
	/**
	 * Get the worksheet automation from the embedded workbook based on the index number  
	 * @param workbookAutomation an OleAutomation that provides access to the functionalities of a Workbook OLE object
	 * @param index an integer that represents the index of the worksheet in the collection of worksheets for the given workbook automation.
	 * @return an OleAutomation to access worksheet object functionalities
	 */
	public static OleAutomation getWorksheetAutomationByIndex(OleAutomation workbookAutomation, int index){
		
		OleAutomation worksheetsAutomation = getWorksheetsAutomation(workbookAutomation);
		if(worksheetsAutomation==null){
			System.out.println("ERROR: Could not retrieve Worksheets automation!!!");
			System.exit(1);
		}	
		
		OleAutomation worksheetAutomation = CollectionsUtils.getItemByIndex(worksheetsAutomation, index, false);
		worksheetsAutomation.dispose();
		
		return worksheetAutomation;
	}	
	
	
	/**
	 * Add a new worksheet to the given workbook
	 * @param workbookAutomation an OleAutomation that provides access to the functionalities of a Workbook OLE object
	 * @return an OleAutomation for the new worksheet 
	 */
	public static OleAutomation addWorksheetAsLast(OleAutomation workbookAutomation){
		
		OleAutomation worksheetsAutomation = getWorksheetsAutomation(workbookAutomation);
		if(worksheetsAutomation==null){
			System.out.println("ERROR: Could not retrieve Worksheets automation!!!");
			System.exit(1);
		}	
		
		int count = CollectionsUtils.getNumberOfObjectsInOleCollection(worksheetsAutomation);
		OleAutomation lastSheet = getWorksheetAutomationByIndex(workbookAutomation, count); 
		
		int[] addMethodIds = worksheetsAutomation.getIDsOfNames(new String[]{"Add", "After"});	
		Variant[] params = new Variant[1];
		params[0] = new Variant(lastSheet);
		int paramsIds[] = Arrays.copyOfRange(addMethodIds, 1, addMethodIds.length);
		Variant result = worksheetsAutomation.invoke(addMethodIds[0], params, paramsIds);
		
		OleAutomation newWorksheet = result.getAutomation();
		for (Variant v : params) {
			v.dispose();
		}
		result.dispose();
		
		return newWorksheet;
	}
	
	
	/**
	 * Protect the structure of the active workbook 
	 * @param workbookAutomation an OleAutomation that provides access to the functionalities of a Workbook OLE object
	 * @return true if operation succeeded, false otherwise
	 */
	public static boolean protectWorkbook(OleAutomation workbookAutomation, boolean structure, boolean windows){
		
		// invoke the "Protect" method for the given workbook
		int[] protectMethodIds = workbookAutomation.getIDsOfNames(new String[]{"Protect", "Structure", "Windows"});
		if (protectMethodIds == null) {
			System.out.println("Method \"Protect\" not found for \"Workbook\" object!");
			return false;
		}else{
			Variant[] args = new Variant[2];
			args[0] = new Variant(structure);
			args[1] = new Variant(windows);
			
			int argsIds[] = Arrays.copyOfRange(protectMethodIds, 1, protectMethodIds.length);
			Variant result = workbookAutomation.invoke(protectMethodIds[0], args, argsIds);	
			
			for (Variant arg: args) {
				arg.dispose();
			}
			
			if(result==null)
				return false;
		
			result.dispose();
		}		
		return true;
	}
	
	
   /**
	 * Unprotect the structure of the active workbook
	 * @param workbookAutomation an OleAutomation that provides access to the functionalities of a Workbook OLE object
	 * @return true if operation succeeded, false otherwise
	 */
	public static boolean unprotectWorkbook(OleAutomation workbookAutomation){
		
		// invoke the "Unprotect" method for the given workbook
		int[] unprotectMethodIds = workbookAutomation.getIDsOfNames(new String[]{"Unprotect"});
		if (unprotectMethodIds == null) {
			System.out.println("Method \"Unprotect\" not found for \"Workbook\" object!");
			return false;
		}else{			
			Variant result = workbookAutomation.invoke(unprotectMethodIds[0]);	
			
			if(result==null)
				return false;
		
			result.dispose();
		}
		return true;
	}

	
	/**
	 * Protect all worksheets that are part of the given workbook
	 * @param workbookAutomation an OleAutomation that provides access to the functionalities of a Workbook OLE object
	 * @return true if operation succeeded, false otherwise
	 */
	public static boolean protectAllWorksheets(OleAutomation workbookAutomation){	
		OleAutomation worksheetsAutomation = getWorksheetsAutomation(workbookAutomation);
		
		int count = CollectionsUtils.getNumberOfObjectsInOleCollection(worksheetsAutomation);
		
		int i; 
		boolean isSuccess=true; 
		for (i = 1; i <= count; i++) {
		
			OleAutomation nextWorksheetAutomation = CollectionsUtils.getItemByIndex(worksheetsAutomation, i, false);					
			if(!WorksheetUtils.protectWorksheet(nextWorksheetAutomation)){
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
				OleAutomation nextWorksheetAutomation =  CollectionsUtils.getItemByIndex(worksheetsAutomation, j, false);
				WorksheetUtils.unprotectWorksheet(nextWorksheetAutomation);
				nextWorksheetAutomation.dispose();
			}
			worksheetsAutomation.dispose();
			return false;
		}
		
		worksheetsAutomation.dispose();
		return true;
	}
	
	
	/**
	 * Close the given workbook. Invoke the "Close" method of the Workbook OLe object. Specify if should save changes.     
	 * @param workbookAutomation an OleAutomation that provides access to the functionalities of the Workbook OLE object 
	 * @param saveChanges if yes, changes made in the documents will be saved. Otherwise, all changes will be discarded. 
	 */
	public static boolean closeEmbeddedWorkbook(OleAutomation workbookAutomation, boolean saveChanges){
		
		if(workbookAutomation==null){
			System.out.println("ERROR: Workbook is null!!!");
			return false;
		}		
		//TODO: implement when saveChanges = true
		int[] closeMethodIds = workbookAutomation.getIDsOfNames(new String[]{"Close","SaveChanges"}); //"Filename"	
		if (closeMethodIds == null){			
			System.out.println("\"Close\" method not found for \"Workbook\" object!");
			return false;
		}	
		
		Variant[] args = new Variant[1]; 
		args[0] = new Variant(saveChanges);
		
		int[] argumentIds = Arrays.copyOfRange(closeMethodIds, 1, closeMethodIds.length); 
		Variant result = workbookAutomation.invoke(closeMethodIds[0], args, argumentIds);
		if(result==null){ // || result.getType() == OLE.VT_EMPTY)
			return false;
		}
		
		result.dispose();
		return true;
	}
}
