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
			System.out.println("\"Name\" variant is null!");
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
	public static void closeEmbeddedWorkbook(OleAutomation workbookAutomation, boolean saveChanges){
		
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
	 *
	 * Get the active worksheet automation using the "ActiveSheet" property. 
	 * @param application an OleAutomation  object that has the "ActiveSheet" property.
	 * @return
	 */
	public static OleAutomation getActiveWorksheetAutomation(OleAutomation automation){
		
		int[] worksheetIds = automation.getIDsOfNames(new String[]{"ActiveSheet"});	
		if (worksheetIds == null){			
			System.out.println("\"ActiveSheet\" property not found for the given OleAutomation object!");
			return null;
		}		
		Variant worksheetVariant = automation.getProperty(worksheetIds[0]);
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
	 * @param sheetName
	 * @return
	 */
	public static OleAutomation getWorksheetAutomationByName(String sheetName){
		
		OleAutomation application = getApplicationAutomation(MainWindow.getInstance().getControlSite());		
		OleAutomation embeddedWorkbook = getEmbeddedWorkbookAutomation(application);
		OleAutomation worksheetsAutomation = getWorksheetsAutomation(embeddedWorkbook);
		
		if(worksheetsAutomation==null){
			System.out.println("ERROR: Could not receive Worksheets automation!!!");
			return null;
		}
		
		OleAutomation sheetAutomation = getItem(worksheetsAutomation, sheetName);	
		worksheetsAutomation.dispose();
		embeddedWorkbook.dispose();
		application.dispose();

		return sheetAutomation;
	}
	
	/**
	 * Get the name of the given worksheet
	 * @param worksheetAutomation an OleAutomation for accessing the worksheet object
	 * @return
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
	 * @param worksheetAutomation an OleAutomation for accessing the worksheet object
	 * @return
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
	 * Get the OleAutomation object for the "Shapes" property of the given worksheet  
	 * @param worksheetAutomation
	 * @return
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
	 * Get the specified range automation. The address of the top left cell and down right cell have to be provided.
	 * The OleAutomation object that is used to retrieve the range has to have the "Range" property. 
	 * @param automation an OleAutomation object that has the "Range" property
	 * @param topLeftCell address of top left cell (e.g., "A1" or "$A$1" )
	 * @param downRightCell address of down right cell (e.g., "C3" or "$C$3" )
	 * @return
	 */
	public static OleAutomation getRangeAutomation(OleAutomation automation, String topLeftCell, String downRightCell){
		
		// get the OleAutomation object for the selected range 
		int[] rangePropertyIds = automation.getIDsOfNames(new String[]{"Range"});
		
		Variant[] args;
		if(downRightCell!=null && downRightCell.compareTo("")!=0){
			args = new Variant[2];
			args[0] = new Variant(topLeftCell);
			args[1] = new Variant(downRightCell);
		}else{
			args = new Variant[1];
			args[0] = new Variant(topLeftCell);
		}
		
		Variant rangeVariant = automation.getProperty(rangePropertyIds[0],args);
		OleAutomation rangeAutomation = rangeVariant.getAutomation();
		for (Variant arg : args) {
			arg.dispose();
		}
		rangeVariant.dispose();
		
		return rangeAutomation;
	}
	
	/**
	 * Get the distance, in points, from the left edge of column A to the left edge of the range.
	 * @param rangeAutomation
	 * @return
	 */
	public static double getRangeLeftPosition(OleAutomation rangeAutomation){

		int[] leftPropertyIds = rangeAutomation.getIDsOfNames(new String[]{"Left"});
		Variant leftVariant=rangeAutomation.getProperty(leftPropertyIds[0]);
		double left = leftVariant.getDouble();
		leftVariant.dispose();
		return left;
	}
	
	/**
	 * Get the distance, in points, from the top edge of row 1 to the top edge of the range
	 * @param rangeAutomation
	 * @return
	 */
	public static double getRangeTopPosition(OleAutomation rangeAutomation){
		
		int[] topPropertyIds = rangeAutomation.getIDsOfNames(new String[]{"Top"});
		Variant topVariant=rangeAutomation.getProperty(topPropertyIds[0]);
		double top = topVariant.getDouble();
		topVariant.dispose();
		
		return top;
	}
	
	/**
	 * Get the height, in units, of the range.
	 * @param rangeAutomation
	 * @return
	 */
	public static double getRangeHeight(OleAutomation rangeAutomation){
		
		int[] heightPropertyIds = rangeAutomation.getIDsOfNames(new String[]{"Height"});
		Variant heightVariant=rangeAutomation.getProperty(heightPropertyIds[0]);
		double height = heightVariant.getDouble();
		heightVariant.dispose();
		
		return height;
	}
	
	
	/**
	 * Get the width, in units, of the range.
	 * @param rangeAutomation
	 * @return
	 */
	public static double getRangeWidth(OleAutomation rangeAutomation){
		
		int[] widthPropertyIds = rangeAutomation.getIDsOfNames(new String[]{"Width"});
		Variant widthVariant=rangeAutomation.getProperty(widthPropertyIds[0]);
		double width = widthVariant.getDouble();
		widthVariant.dispose();
		
		return width;
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
