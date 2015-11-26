/**
 * 
 */
package de.tudresden.annotator.main;

import org.eclipse.swt.ole.win32.OleAutomation;
import org.eclipse.swt.ole.win32.Variant;

/**
 * @author Elvis
 *
 */
public class OleActionHelper {
	
	protected static boolean getListOfCommandBars(OleAutomation excelApplication) {
		
		int[] commandBarsObjectIds = excelApplication.getIDsOfNames(new String[]{"CommandBars"});
		if (commandBarsObjectIds == null) {
			System.out.println("Property \"CommandBars\" of \"Application\" OLE Object is null!");
			return false;
		}
		
		Variant commandBarsVariant =  excelApplication.getProperty(commandBarsObjectIds[0]);	
		if(commandBarsVariant == null){
			System.out.println("\"CommandBars\" variant is null!");
			return false;		
		}
		OleAutomation commandBarsAutomation = commandBarsVariant.getAutomation();
		commandBarsVariant.dispose();
		
		int[] countProperyIds = commandBarsAutomation.getIDsOfNames(new String[]{"Count"});
		if(countProperyIds == null){
			System.out.println("Property \"Count\" of \"CommandBars\" OLE object is null!");
			return false;
		}
				
		Variant countPropertyVariant =  commandBarsAutomation.getProperty(countProperyIds[0]);
		if(countPropertyVariant == null){
			System.out.println("\"Count\" variant is null!");
			return false;
		}				
				
		int count = countPropertyVariant.getInt();
		countPropertyVariant.dispose();
		
		int[] itemPropertyIds = commandBarsAutomation.getIDsOfNames(new String[]{"Item"});
		if(itemPropertyIds == null){
			System.out.println("Property \"Item\" of \"CommandBars\" is not found!");
			return false;
		}

		System.out.println("\nlist of command bars:\n".toUpperCase());
		for (int i = 1; i <= count; i++) {
			Variant[] args = new Variant[1];
			args[0] = new Variant(i);		
			Variant nextCommandBarVariant = commandBarsAutomation.getProperty(itemPropertyIds[0],args);	
			
			OleAutomation nextCommandBarAutomation = nextCommandBarVariant.getAutomation();
			System.out.println(getCommandBarName(nextCommandBarAutomation));
			
			nextCommandBarVariant.dispose();
			nextCommandBarAutomation.dispose();
			args[0].dispose();
		}
		return true;	
	}
	
	protected static String getCommandBarName(OleAutomation commandBarAutomation){
		
		int[] namePropertyIds = commandBarAutomation.getIDsOfNames(new String[]{"Name"});
		if(namePropertyIds == null){
			System.out.println("Property \"Name\" of \"CommandBar\" is not found!");
			return null;
		}
		Variant nameVariant = commandBarAutomation.getProperty(namePropertyIds[0]);
		if(nameVariant == null){
			System.out.println("\"Name\" variant is null!");
			return null;
		}
		String name = nameVariant.getString();
		nameVariant.dispose();
		return name;
	}
	
	
	/**
	 * Hide the formula bar from Excel UI
	 * @param excelApplication
	 * @return
	 */
	protected static boolean hideFormulaBar(OleAutomation excelApplication){
			
		int[] displayFormulaBarIds = excelApplication.getIDsOfNames(new String[]{"DisplayFormulaBar"});
		
		//Variant  displayFormulaBarVariant =  application.getProperty(displayFormulaBarIds[0]);	
		//System.out.println("Initial value DisplayFormulaBar: "+displayFormulaBarVariant);
		
		Boolean isUpdated = excelApplication.setProperty(displayFormulaBarIds[0],new Variant(false));
		//System.out.println("Property is updated? "+isUpdated);
		
		//displayFormulaBarVariant =  application.getProperty(displayFormulaBarIds[0]);
		//System.out.println("New value DisplayFormulaBar: "+displayFormulaBarVariant);
		
		//displayFormulaBarVariant.dispose();
		
		return isUpdated;
	}
	
	
	/**
	 * Set whether to display or not floating menus on right click  
	 * 
	 * @param excelApplication
	 * @return
	 */
	protected static boolean setShowMenuFloaties(OleAutomation excelApplication, boolean option){
		
		int[] showMenuFloatiesIds = excelApplication.getIDsOfNames(new String[]{"ShowMenuFloaties"});	
		if (showMenuFloatiesIds == null){			
			System.out.println("\"ShowMenuFloaties\" property not found for \"Application\" OLE object!");
			return false;
		}	
		
		boolean result = excelApplication.setProperty(showMenuFloatiesIds[0],new Variant(option));
		System.out.println("Set \"ShowMenuFloaties\" was successful? "+result);
		return result;
	}
	
	
	/**
	 * Set whether to display or not the developer tab
	 * 
	 * @param excelApplication
	 * @return
	 */
	protected static boolean setShowDevTools(OleAutomation excelApplication, boolean option){
		
		int[] showDevToolsIds = excelApplication.getIDsOfNames(new String[]{"ShowDevTools"});	
		if (showDevToolsIds == null){			
			System.out.println("\"ShowDevTools\" property not found for \"Application\" OLE object!");
			return false;
		}	
		
		boolean result = excelApplication.setProperty(showDevToolsIds[0],new Variant(option));
		System.out.println("Set \"ShowDevTools\" was successful? "+result);
		return result;
	}
	
	
	/**
	 * Set whether to display or not tooltips
	 * 
	 * @param excelApplication
	 * @return
	 */
	protected static boolean setShowToolTips(OleAutomation excelApplication, boolean option){
		
		int[] showToolTipsIds = excelApplication.getIDsOfNames(new String[]{"ShowToolTips"});	
		if (showToolTipsIds == null){			
			System.out.println("\"ShowToolTips\" property not found for \"Application\" OLE object!");
			return false;
		}	
		
		boolean result = excelApplication.setProperty(showToolTipsIds[0],new Variant(option));
		System.out.println("Set \"ShowToolTips\" was successful? "+result);
		return result;
	}
	
	
	/**
	 * Get the command bar automation by its name
 	 * @param application
	 * @return
	 */
	protected static OleAutomation getCommandBarByName(OleAutomation application, String commandBarName) {
		
		int[] commandBarsPropertyIds = application.getIDsOfNames(new String[]{"CommandBars"});
		if (commandBarsPropertyIds == null) {
			System.out.println("Property \"CommandBars\" of \"Application\" OLE Object is null!");
			return null;
		}
		
		Variant commandBarsVariant =  application.getProperty(commandBarsPropertyIds[0]);	
		if(commandBarsVariant == null){
			System.out.println("\"CommandBars\" variant is null!");
			return null;		
		}
		OleAutomation commandBarsAutomation = commandBarsVariant.getAutomation();
		commandBarsVariant.dispose();
			
		int[] itemPropertyIds = commandBarsAutomation.getIDsOfNames(new String[]{"Item"});
		if(itemPropertyIds == null){
			System.out.println("Property \"Item\" of \"CommandBars\" OLE object not found!");
			return null;
		}

		Variant[] parameters = new Variant[1];
		parameters[0] = new Variant(commandBarName);
		Variant cbVariant = commandBarsAutomation.getProperty(itemPropertyIds[0],parameters);
		parameters[0].dispose();
		
		if(cbVariant==null){
			System.out.println("There is no CommandBar named \""+commandBarName+"\"");
			return null;
		}
		OleAutomation commandBarAutomation = cbVariant.getAutomation();
		cbVariant.dispose();

		return commandBarAutomation;
	}
	
	
	protected static OleAutomation getControls(OleAutomation commandBarAutomation){
		
		int[] controlsPropertyIds = commandBarAutomation.getIDsOfNames(new String[]{"Controls"});
		Variant controlsVariant = commandBarAutomation.getProperty(controlsPropertyIds[0]);
		OleAutomation contolsAutomation = controlsVariant.getAutomation();
		controlsVariant.dispose();
		
		return contolsAutomation;
	}
	
	
	protected static boolean hideAllControls(OleAutomation commandBarContols){

		//make existing controls not visible
		int[] itemPropertyIds = commandBarContols.getIDsOfNames(new String[]{"Item"});
	
		Variant[] parameters = new Variant[1];
		parameters[0] = new Variant(1);
		Variant controlItemVariant = commandBarContols.getProperty(itemPropertyIds[0],parameters);
		parameters[0].dispose();
		
		int i=1;
		while (controlItemVariant!=null) {			
			OleAutomation controlItemAutomation = controlItemVariant.getAutomation();
			int[] visiblePropertyIds = controlItemAutomation.getIDsOfNames(new String[]{"Visible"});
			controlItemAutomation.setProperty(visiblePropertyIds[0],new Variant(false));
			parameters[0] = new Variant(i++);
			controlItemVariant.dispose();
			controlItemVariant = commandBarContols.getProperty(itemPropertyIds[0],parameters);
			parameters[0].dispose();
		}

		return true;
	}
}
