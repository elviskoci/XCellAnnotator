/**
 * 
 */
package de.tudresden.annotator.main;

import java.util.Arrays;

import org.eclipse.swt.ole.win32.OleAutomation;
import org.eclipse.swt.ole.win32.OleControlSite;
import org.eclipse.swt.ole.win32.Variant;

/**
 * @author Elvis
 *
 */
public class OleInterfaceModifier {
		
	/**
	 * Get Excel application as an OLE object automation
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
	 * Get OLE automation for the active workbook 
	 * 
	 * @param application
	 * @return
	 */
	public static OleAutomation getActiveWorkbook(OleAutomation application){
		
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
	 * Protect the structure of the active workbook 
	 * @param application
	 * @return
	 */
	public static boolean protectWorkbook(OleAutomation workbookAutomation){
		
		if (workbookAutomation==null)
			return false;
		
		// invoke the "Protect" method for the active workbook
		int[] protectMethodIds = workbookAutomation.getIDsOfNames(new String[]{"Protect", "Structure", "Windows"});
		if (protectMethodIds == null) {
			System.out.println("Method \"Protect\" not found for \"Workbook\" object!");
			return false;
		}else{
			Variant[] args = new Variant[2];
			args[0] = new Variant(true);
			args[1] = new Variant(false);
			
			Variant result = workbookAutomation.invoke(protectMethodIds[0],args,Arrays.copyOfRange(protectMethodIds, 1, protectMethodIds.length));	
//			System.out.println("Result of Workbook.Protect(): "+result);
			if(result==null)
				return false;
			
			result.dispose();
			for (Variant arg: args) {
				arg.dispose();
			}
		}
		
		return true;
	}
	
	
	/**
	 * Unprotect the structure of the active workbook
	 * @param application
	 * @return
	 */
	public static boolean unprotectActiveWorkbook(OleAutomation workbookAutomation){
		
		if (workbookAutomation==null)
			return false;
				
		// invoke the "Protect" method for the active workbook
		int[] unprotectMethodIds = workbookAutomation.getIDsOfNames(new String[]{"Unprotect"});
		if (unprotectMethodIds == null) {
			System.out.println("Method \"Unprotect\" not found for \"Workbook\" object!");
			return false;
		}else{
			Variant[] args = new Variant[1];
			args[0] = new Variant();
			
			Variant result = workbookAutomation.invoke(unprotectMethodIds[0],args);	
//			System.out.println("Result of Workbook.Unprotect(): "+result);
			if(result==null)
				return false;
			
			result.dispose();
			args[0].dispose();
		}
		
		return true;
	}
	

	/**
	 * Protect all worksheet that are part of the given workbook
	 * @param application
	 * @return
	 */
	public static boolean protectAllWorksheets(OleAutomation workbookAutomation){
		
		// mark each worksheet as protected 
		int[] worksheetsObjectIds = workbookAutomation.getIDsOfNames(new String[]{"Worksheets"});
		if (worksheetsObjectIds == null) {
			System.out.println("Property \"Worksheets\" of \"Workbook\" OLE Object is null!");
			return false;
		}
		
		Variant worksheetsVariant =  workbookAutomation.getProperty(worksheetsObjectIds[0]);	
		if(worksheetsVariant == null){
			System.out.println("\"Worksheets\" variant is null!");
			return false;		
		}
		OleAutomation worksheetsAutomation = worksheetsVariant.getAutomation();
		worksheetsVariant.dispose();
		
		int[] countProperyIds = worksheetsAutomation.getIDsOfNames(new String[]{"Count"});
		if(countProperyIds == null){
			System.out.println("Property \"Count\" of \"Worksheets\" OLE object is null!");
			return false;
		}
				
		Variant countPropertyVariant =  worksheetsAutomation.getProperty(countProperyIds[0]);
		if(countPropertyVariant == null){
			System.out.println("\"Count\" variant is null!");
			return false;
		}				
				
		int count = countPropertyVariant.getInt();
		countPropertyVariant.dispose();
		
		int[] itemPropertyIds = worksheetsAutomation.getIDsOfNames(new String[]{"Item"});
		if(itemPropertyIds == null){
			System.out.println("Property \"Item\" of \"Worksheets\" OLE object not found!");
			return false;
		}
		
		int i; 
		boolean isSuccess=true; 
		for (i = 1; i <= count; i++) {
			Variant[] args = new Variant[1];
			args[0] = new Variant(i);		
			
			Variant nextWorkbookVariant = worksheetsAutomation.getProperty(itemPropertyIds[0],args);	
			OleAutomation nextWorkbookAutomation = nextWorkbookVariant.getAutomation();	
			if(!protectWorksheet(nextWorkbookAutomation)){
				System.out.println("ERROR: Could not protect one of the workbooks!");
				isSuccess=false;
				
			}
			
			nextWorkbookVariant.dispose();
			nextWorkbookAutomation.dispose();
			args[0].dispose();
			
			if(!isSuccess){
				break;
			}
		}	
		
		if(!isSuccess){
			for(int j=1; j<i;j++){
				Variant[] args = new Variant[1];
				args[0] = new Variant(j);		
				
				Variant nextWorkbookVariant = worksheetsAutomation.getProperty(itemPropertyIds[0],args);
				OleAutomation nextWorkbookAutomation = nextWorkbookVariant.getAutomation();
				unprotectWorksheet(nextWorkbookAutomation);
				
				nextWorkbookVariant.dispose();
				nextWorkbookAutomation.dispose();
				args[0].dispose();
			}
			return false;
		}
		
		return true;
	}
	
	
	/**
	 * Protect the data, formating, and structure of the specified worksheet
	 * @param worksheetAutomation
	 * @return
	 */
	public static boolean protectWorksheet(OleAutomation worksheetAutomation){
		
		// get the id of the "Protect" method and the considered parameters
		// you can find the documentation of this OLE method here https://msdn.microsoft.com/EN-US/library/ff840611.aspx
		int[] protectMethodIds = worksheetAutomation.getIDsOfNames(new String[]{"Protect", "DrawingObjects", "Contents", "Scenarios", 
				"AllowFormattingCells", "AllowFormattingColumns", "AllowFormattingRows", "AllowInsertingColumns", "AllowInsertingRows", 
				"AllowInsertingHyperlinks", "AllowDeletingColumns", "AllowDeletingRows", "AllowSorting", "AllowFiltering", "AllowUsingPivotTables" });
		
		if (protectMethodIds == null) {
			System.out.println("Method \"Protect\" of \"Worksheet\" OLE Object is not found!");
			return false;
		}else{
			Variant[] args = new Variant[14];
			args[0] = new Variant(true);
			args[1] = new Variant(true);
			args[2] = new Variant(true);
			args[3] = new Variant(false);
			args[4] = new Variant(false);
			args[5] = new Variant(false);
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
	 * 
	 * @param worksheetAutomation
	 * @return
	 */
	public static boolean unprotectWorksheet(OleAutomation worksheetAutomation){
		
		// get the id of the "Unprotect" method for worksheet OLE object 
		int[] unprotectMethodIds = worksheetAutomation.getIDsOfNames(new String[]{"Unprotect"});
		if(unprotectMethodIds==null){
			System.out.println("Method \"Unprotect\" of \"Worksheet\" OLE Object is not found!");
			return false;
		}
		
		// call the unprotect method  
		Variant result = worksheetAutomation.invoke(unprotectMethodIds[0]);

		if(result==Variant.NULL){
			result.dispose();
			return false;
		}	
		
		result.dispose();
		return true;
	}
	
	/**
	 * Hide Ribbon from Excel GUI
	 * @param application
	 * @return
	 */
	public static boolean hideRibbon(OleAutomation application){
		
		int[] ee4mIds = application.getIDsOfNames(new String[]{"ExecuteExcel4Macro"});
		
		Variant[] parameters = new Variant[1];
	    parameters[0] = new Variant("SHOW.TOOLBAR(\"Ribbon\",False)");
	    
	    Variant result = application.invoke(ee4mIds[0],parameters);
//	    System.out.println("\nThe result of ExecuteExcel4Macro method invocation: "+result);
	   
	    boolean isSuccess = false;
	    if(result!=null)
	    	isSuccess = true;
	    
	    parameters[0].dispose();
	    result.dispose();	
	    
	    return isSuccess;
	}
	
	
	/**
	 * Create a custom "Cell" command bar. All the exiting controls (menu items) will be hidden, and new ones will be added.
	 * @param application
	 */
	public static void createCustomCellCommandBar(OleAutomation application){
		
		// get the "Cell" command bar automation
		OleAutomation cellCBAutomation = CommandBarsHelper.getCommandBarByName(application,"cell");		
		if(cellCBAutomation==null)
			return;
		
		// get CommandBarsControls object automation.
		OleAutomation contolsAutomation = CommandBarsHelper.getCommandBarControls(cellCBAutomation);
		cellCBAutomation.dispose();
		
		// temporary delete the (menu) items in the "Cell" command bar. This controls will appear again when Excel application is started  
		CommandBarsHelper.deleteControlsTemporary(contolsAutomation);
		
		// add new CommandBarPopup control
		int[] addMethodIds = contolsAutomation.getIDsOfNames(new String[]{"Add", "Type", "Before"});
		
		Variant[] args = new Variant[addMethodIds.length-1];
		args[0] = new Variant(10);
		args[1] = new Variant(1);
		
		Variant myControlItemVariant = contolsAutomation.invoke(addMethodIds[0],args,Arrays.copyOfRange(addMethodIds, 1, addMethodIds.length));
		OleAutomation myPopUpControl = myControlItemVariant.getAutomation();
		myControlItemVariant.dispose();	
		for (Variant arg : args) {
			arg.dispose();
		}
		
		// set the properties of the control
		int[] captionProperyIds = myPopUpControl.getIDsOfNames(new String[]{"Caption"});
		int[] tagProperyIds = myPopUpControl.getIDsOfNames(new String[]{"Tag"});
		myPopUpControl.setProperty(captionProperyIds[0], new Variant("Annotate as"));
		myPopUpControl.setProperty(tagProperyIds[0], new Variant("annotation_controls"));
		
		// add sub-controls (sub-menus) 
		OleAutomation mySubContolsAutomation = CommandBarsHelper.getCommandBarControls(myPopUpControl);
		addMethodIds = contolsAutomation.getIDsOfNames(new String[]{"Add", "Type"});
		args = new Variant[addMethodIds.length-1];
		args[0] = new Variant(1);
		
		String[] captions = new String[]{"Table","Metadata","Header","Attributes","Data"};
		Variant[] mySubControlVariant = new Variant[captions.length];
		OleAutomation mySubControl;
		for (int i = 0; i < captions.length; i++) {
			mySubControlVariant[i] = mySubContolsAutomation.invoke(addMethodIds[0],args,Arrays.copyOfRange(addMethodIds, 1, addMethodIds.length));		
			
			mySubControl= mySubControlVariant[i].getAutomation(); 
			int[] captionPropetyIds = mySubControl.getIDsOfNames(new String[]{"Caption"});
			mySubControl.setProperty(captionPropetyIds[0], new Variant(captions[i]));
			
//			int[] onActionPropertyIds = mySubControl.getIDsOfNames(new String[]{"OnAction"});
//			mySubControl.setProperty(onActionPropertyIds[0], new Variant("MsgBox \"You annotated as ...\""));
//			
//			MainWindow.getInstance().getControlSite().addPropertyListener(onActionPropertyIds[0], new OleListener() {			
//				@Override
//				public void handleEvent(OleEvent event) {
//					System.out.println("Event Captured!!!");
//				}
//			});
			
			mySubControlVariant[i].dispose();
			mySubControl.dispose();
		}
		contolsAutomation.dispose();
		mySubContolsAutomation.dispose();
		for (Variant arg : args) {
			arg.dispose();
		}	
	}
	
	
	/**
	 * Undo changes done to the "Cell" commandbar during the current session
	 * @param application
	 */
	public static void undoChangesToCellCommandBar(OleAutomation application){
		
	    // get the "Cell" command bar automation
		OleAutomation cellCBAutomation = CommandBarsHelper.getCommandBarByName(application,"cell");		
		if(cellCBAutomation==null)
			return;
		
		// get CommandBarsControls object automation.
		OleAutomation contolsAutomation = CommandBarsHelper.getCommandBarControls(cellCBAutomation);
		cellCBAutomation.dispose();
		
		// show (Make visible) all the control (menu) items in the "Cell" command bar
		CommandBarsHelper.setVisibilityOfControls(contolsAutomation, true);
		// delete custom controls, created during the current session of the application 
		CommandBarsHelper.deleteCustomControlsByTag(contolsAutomation,"annotation_controls");		
	}
}
