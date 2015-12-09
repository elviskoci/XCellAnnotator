/**
 * 
 */
package de.tudresden.annotator.oleutils;

import org.eclipse.swt.ole.win32.OleAutomation;
import org.eclipse.swt.ole.win32.OleControlSite;
import org.eclipse.swt.ole.win32.Variant;

/**
 * @author Elvis Koci
 */
public class ApplicationUtils {
	
	/**
	 * Get Excel application as an OleAutomation object
	 * @param controlSite the OleControlSide for the embedded spreadsheet file 
	 * @return an OleAutomation that provides access to the functionalities of the (Excel) Application OLE object 
	 */
	public static OleAutomation getApplicationAutomation(OleControlSite controlSite){
		
	    OleAutomation excelClient = new OleAutomation(controlSite);
		int[] dispIDs = excelClient.getIDsOfNames(new String[] {"Application"});
		
		if(dispIDs==null){	
			System.out.println("\"Application\" object not found!");
			return null;
		}
		
		Variant pVarResult = excelClient.getProperty(dispIDs[0]);
		if(pVarResult==null){	
			System.out.println("\"Application\" object is null!");
			return null;
		}
		
		OleAutomation application = pVarResult.getAutomation();
		
		pVarResult.dispose();
		excelClient.dispose();
		
		return application;
	}
	

	/**
	 * Get OleAutomation for the active workbook using the "ActiveWorkbook" property. 
	 * Excel application considers the workbook which has the focus to be the "active" one.
	 *  
	 * @param application an OleAutomation that provides access to the functionalities of the (Excel) Application OLE object
	 * @return an OleAutomation that provides access to the functionalities of the Active Workbook OLE object
	 */
	public static OleAutomation getActiveWorkbookAutomation(OleAutomation application){
		
		int[] workbookIds = application.getIDsOfNames(new String[]{"ActiveWorkbook"});	
		if (workbookIds == null){			
			System.out.println("\"ActiveWorkbook\" property not found for \"Application\" object!");
			return null;
		}		
		Variant workbookVariant = application.getProperty(workbookIds[0]);
		if (workbookVariant == null) {
			System.out.println("Workbook variant is null!");
			return null;
		}		
		OleAutomation workbookAutomation =  workbookVariant.getAutomation();
		workbookVariant.dispose();
		
		return workbookAutomation;
	}
	
	
	/**
	 * Get the OleAutomation object for the embedded workbook  
	 * @param application an OleAutomation that provides access to the functionalities of the (Excel) Application OLE object
	 * @param workbookName the name of the embedded workbook
	 * @return an OleAutomation that provides access to the functionalities of the Embedded Workbook OLE object
	 */
	public static OleAutomation getEmbeddedWorkbookAutomation(OleAutomation application, String workbookName){
		
		int[] workbooksIds = application.getIDsOfNames(new String[]{"Workbooks"});	
		if (workbooksIds == null){			
			System.out.println("\"Workbooks\" property not found for \"Application\" object!");
			return null;
		}		
		
		Variant workbooksVariant = application.getProperty(workbooksIds[0]);
		if (workbooksVariant == null) {
			System.out.println("Workbooks variant is null!");
			return null;
		}
		
		OleAutomation workbooksAutomation = workbooksVariant.getAutomation();
		workbooksVariant.dispose();
			
		OleAutomation embeddedWorkbook = CollectionsUtils.getItemByName(workbooksAutomation, workbookName);
		workbooksAutomation.dispose();
		
		return embeddedWorkbook;	
	}
	
	
	/**
	 * Get the workbook OleAutomation using the "ThisWorkbook" property  
	 * @param application an OleAutomation that provides access to the functionalities of the (Excel) Application OLE object
	 * @return an OleAutomation that provides access to the functionalities of the This Workbook OLE object
	 */
	public static OleAutomation getThisWorkbookAutomation(OleAutomation application){
		
		int[] thisWorkbookIds = application.getIDsOfNames(new String[]{"ThisWorkbook"});	
		if (thisWorkbookIds == null){			
			System.out.println("\"ThisWorkbook\" property not found for \"Application\" object!");
			return null;
		}		
		
		Variant thisWorkbookVariant = application.getProperty(thisWorkbookIds[0]);
		if (thisWorkbookVariant == null) {
			System.out.println("ThisWorkbook variant is null!");
			return null;
		}
		
		OleAutomation workbookAutomation = thisWorkbookVariant.getAutomation();
		thisWorkbookVariant.dispose();
		
		return workbookAutomation;
	}
		
	/**
	 * Get the Worksheets automation from the "active" Workbook
	 * @param application an OleAutomation that provides access to the functionalities of the (Excel) Application OLE object 
	 * @return an OleAutomation that provides access to the functionalities of the Worksheets OLE object
	 */
	public static OleAutomation getWorksheetsAutomation(OleAutomation application){
		
		// get ID of Worksheets property
		int[] worksheetsObjectIds = application.getIDsOfNames(new String[]{"Worksheets"});
		if (worksheetsObjectIds == null) {
			System.out.println("Property \"Worksheets\" was not found for the given Application OLE object!");
			return null;
		}
		
		// get property using the ID 
		Variant worksheetsVariant =  application.getProperty(worksheetsObjectIds[0]);	
		if(worksheetsVariant == null){
			System.out.println("\"Worksheets\" variant is null!");
			return null;		
		}
		
		// get worksheets automation from the variant
		OleAutomation worksheetsAutomation = worksheetsVariant.getAutomation();
		worksheetsVariant.dispose();
		
		return worksheetsAutomation;
	}
	
	/**
	 * Get the active worksheet automation using the "ActiveSheet" property. 
	 * @param applicationAutomation an OleAutomation that provides access to the functionalities of the Excel (Application) OLE object
	 * @return an OleAutomation for the ActiveWorksheet
	 */
	public static OleAutomation getActiveWorksheetAutomation(OleAutomation applicationAutomation){
		
		int[] worksheetIds = applicationAutomation.getIDsOfNames(new String[]{"ActiveSheet"});	
		if (worksheetIds == null){			
			System.out.println("\"ActiveSheet\" property not found for the given OleAutomation object!");
			return null;
		}		
		Variant worksheetVariant = applicationAutomation.getProperty(worksheetIds[0]);
		if (worksheetVariant == null) {
			System.out.println("Workbook variant is null!");
			return null;
		}		
		OleAutomation worksheetAutomation = worksheetVariant.getAutomation();
		worksheetVariant.dispose();
		
		return worksheetAutomation;
	}
	
	/**
	 * Get the specified range automation from the active worksheet. The address of the top left cell and down right cell have to be provided.
	 * The Application OleAutomation object will retrieve the range from the worksheet that is the "active" one at that moment.
	 * @param applicationAutomation an OleAutomation object for accessing the (Excel) Application OLE object
	 * @param topLeftCell address of top left cell (e.g., "A1" or "$A$1" )
	 * @param downRightCell address of down right cell (e.g., "C3" or "$C$3" )
	 * @return
	 */
	public static OleAutomation getRangeAutomation(OleAutomation applicationAutomation, String topLeftCell, String downRightCell){
		
		// get the OleAutomation object for the selected range 
		int[] rangePropertyIds = applicationAutomation.getIDsOfNames(new String[]{"Range"});
		
		Variant[] args;
		if(downRightCell!=null && downRightCell.length()>1){
			args = new Variant[2];
			args[0] = new Variant(topLeftCell);
			args[1] = new Variant(downRightCell);
		}else{
			args = new Variant[1];
			args[0] = new Variant(topLeftCell);
		}
		
		Variant rangeVariant = applicationAutomation.getProperty(rangePropertyIds[0],args);
		OleAutomation rangeAutomation = rangeVariant.getAutomation();
		for (Variant arg : args) {
			arg.dispose();
		}
		rangeVariant.dispose();
		
		return rangeAutomation;
	}
	
	
	/**
	 * Hide Ribbon from Excel GUI
	 * @param application an OleAutomation that provides access to the functionalities of the (Excel) Application OLE object
	 * @return
	 */
	public static boolean hideRibbon(OleAutomation application){
		
		int[] ee4mIds = application.getIDsOfNames(new String[]{"ExecuteExcel4Macro"});
		
		Variant[] parameters = new Variant[1];
	    parameters[0] = new Variant("SHOW.TOOLBAR(\"Ribbon\",False)");
	    
	    Variant result = application.invoke(ee4mIds[0], parameters);
	   
	    boolean isSuccess = false;
	    if(result!=null){
	    	isSuccess = true;
	    	result.dispose();
	    }
	    parameters[0].dispose();
	    
	    return isSuccess;
	}
	
	
	/**
	 * Quit Excel application. Use the given Application OleAutomation to invoke the "Quit" method. 
	 * @param application an OleAutomation that provides access to the functionalities of the (Excel) Application OLE object
	 */
	public static boolean quitExcelApplication(OleAutomation application){
		
		if(application==null){
			System.out.println("ERROR: Application is null!!!");
			return false;
		}
			
		int[] quitMethodIds = application.getIDsOfNames(new String[]{"Quit"});
		if (quitMethodIds == null){			
			System.out.println("\"Quit\" method not found for \"Application\" object!");
			return false;
		}	
		
		Variant result = application.invoke(quitMethodIds[0]);
		if(result==null){ 
			return false;
		}
		
		result.dispose();
		return true;
	}
}
