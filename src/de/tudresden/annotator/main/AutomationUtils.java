package de.tudresden.annotator.main;

import java.util.Arrays;

import org.eclipse.swt.ole.win32.OleAutomation;
import org.eclipse.swt.ole.win32.OleControlSite;
import org.eclipse.swt.ole.win32.Variant;

public class AutomationUtils {
	
	/**
	 * Get Excel application as an OleAutomation object
	 * @param controlSite
	 * @return
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
	 * Quit Excel application
	 * @param application
	 */
	public static void quitExcelApplication(OleAutomation application){
		
		if(application==null){
			System.out.println("ERROR: Application is null!!!");
			return;
		}
			
		int[] quitMethodIds = application.getIDsOfNames(new String[]{"Quit"});
		if (quitMethodIds == null){			
			System.out.println("\"Quit\" method not found for \"Application\" object!");
			return;
		}	
		
		Variant result = application.invoke(quitMethodIds[0]);
		System.out.println(result);
	}
	
	
	/**
	 * Get the Worksheets automation
	 * @param automation an OleAutomation object that has the "Worksheets" property. 
	 * @return
	 */
	public static OleAutomation getWorksheetsAutomation(OleAutomation automation){
		
		// get ID of Worksheets property
		int[] worksheetsObjectIds = automation.getIDsOfNames(new String[]{"Worksheets"});
		if (worksheetsObjectIds == null) {
			System.out.println("Property \"Worksheets\" was not found for the given OLE object!");
			return null;
		}
		
		// get property using the ID 
		Variant worksheetsVariant =  automation.getProperty(worksheetsObjectIds[0]);	
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
	 *
	 * Get OleAutomation for the active workbook using the "ActiveWorkbook" property. 
	 * Excel application considers the workbook which has the focus to be the "active" one.
	 *  
	 * @param application
	 * @return
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
	 * @param application
	 * @param workbookName
	 * @return
	 */
	public static OleAutomation getEmbeddedWorkbookAutomation(OleAutomation application){
		
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
			
		String workbookName = MainWindow.getInstance().getEmbeddedWorkbookName();
		OleAutomation embeddedWorkbook = getItem(workbooksAutomation, workbookName);
		workbooksAutomation.dispose();
		
		return embeddedWorkbook;	
	}
	
	
	/**
	 * Get the workbook OleAutomation using the "ThisWorkbook" property  
	 * @param application
	 * @return
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
	 * Get the name of the given workbook
	 * @param workbookAutomation
	 * @return
	 */
	public static String getWorkbookName(OleAutomation workbookAutomation){
		
		int[] namePropertyIds = workbookAutomation.getIDsOfNames(new String[]{"Name"});	
		if (namePropertyIds == null){			
			System.out.println("\"Name\" property not found for \"Workbook\" object!");
			return null;
		}		
		
		Variant nameVariant = workbookAutomation.getProperty(namePropertyIds[0]);
		if (nameVariant == null) {
			System.out.println("Workbooks variant is null!");
			return null;
		}
		
		String workbookName = nameVariant.getString();
		nameVariant.dispose();
		
		return workbookName;
	}
	
	/**
	 * Close the embedded workbook 
	 * 
	 * @param workbookAutomation
	 * @param saveChanges
	 */
	public static void closeEmbeddedWorksheet(OleAutomation workbookAutomation, boolean saveChanges){
		
		if(workbookAutomation==null){
			System.out.println("ERROR: Workbook is null!!!");
			return;
		}		
		
		int[] closeMethodIds = workbookAutomation.getIDsOfNames(new String[]{"Close","SaveChanges"}); //"Filename"	
		if (closeMethodIds == null){			
			System.out.println("\"Close\" method not found for \"Workbook\" object!");
			return;
		}	
		
		Variant[] args = new Variant[1]; 
		args[0] = new Variant(saveChanges);
		//args[1] = new Variant(MainWindow.getInstance().getEmbeddedWorkbookName());
		
		int[] argumentIds = Arrays.copyOfRange(closeMethodIds, 1, closeMethodIds.length); 
		workbookAutomation.invoke(closeMethodIds[0], args, argumentIds);
	}
	
	
	/**
	 * Get the worksheet automation from the embedded workbook based on the given index  
	 * @param workbookName
	 * @param index
	 * @return
	 */
	public static OleAutomation getWorksheetAutomationByIndex(String index){
		
		OleAutomation application = getApplicationAutomation(MainWindow.getInstance().getControlSite());		
		OleAutomation embeddedWorkbook = getEmbeddedWorkbookAutomation(application);
		OleAutomation worksheetsAutomation = getWorksheetsAutomation(embeddedWorkbook);
		
		if(worksheetsAutomation==null){
			System.out.println("ERROR: Could not receive Worksheets automation!!!");
			return null;
		}
		
		OleAutomation sheetAutomation = getItem(worksheetsAutomation, index);	
		worksheetsAutomation.dispose();
		embeddedWorkbook.dispose();
		application.dispose();

		return sheetAutomation;
	}
	
	/**
	 * Get the item having the specified index from a OleAutomation object. The latter is a collection of OLE Objects. 
	 * This method will fail if the OleAutomation does not have the "Item" property.
	 * @param automation
	 * @param itemName a string that represents the name of the item.
	 * @return
	 */
	public static OleAutomation getItem(OleAutomation automation, String itemName){
		
		int[] itemPropertyIds = automation.getIDsOfNames(new String[]{"Item"});
		if(itemPropertyIds == null){
			System.out.println("Property \"Item\" not found for the give Ole object");
			return null;
		}
		
		Variant args[] = new Variant[1];
		args[0] =  new Variant(itemName);
		
		Variant itemVariant = automation.getProperty(itemPropertyIds[0],args);
		OleAutomation itemAutomation = itemVariant.getAutomation();
		
		args[0].dispose();
		itemVariant.dispose();
		
		return itemAutomation;
	}
	
	
	/**
	 * Get the item having the specified index from a OleAutomation object. The latter is a collection of OLE Objects. 
	 * This method will fail if the OleAutomation does not have the "Item" property.
	 * @param automation
	 * @param index an integer that represents the index of the item in the collection. 
	 * @return
	 */
	public static OleAutomation getItem(OleAutomation automation, int index){
		
		int[] itemPropertyIds = automation.getIDsOfNames(new String[]{"Item"});
		if(itemPropertyIds == null){
			System.out.println("Property \"Item\" not found for the give Ole object");
			return null;
		}
		
		Variant args[] = new Variant[1];
		args[0] =  new Variant(index);
		
		Variant itemVariant = automation.getProperty(itemPropertyIds[0],args);
		OleAutomation itemAutomation = itemVariant.getAutomation();
		
		args[0].dispose();
		itemVariant.dispose();
		
		return itemAutomation;
	}
	
	
	/**
	 * Get the number of items in OleAutomation that is (represents) a collection of OLE objects.
	 * This methods will fail if the given OleAutomation does not have the "Count" property.  
	 * @param automation
	 * @return
	 */
	public static int getNumberOfObjectsInOleCollection(OleAutomation automation){
		
		int[] countProperyIds = automation.getIDsOfNames(new String[]{"Count"});
		if(countProperyIds == null){
			System.out.println("Property \"Count\" not found for the given OleAutomation object!");
			return -1;
		}
				
		Variant countPropertyVariant =  automation.getProperty(countProperyIds[0]);
		if(countPropertyVariant == null){
			System.out.println("\"Count\" variant is null!");
			return -1;
		}				
				
		int count = countPropertyVariant.getInt();
		countPropertyVariant.dispose();
		
		return count;
	}

}
