/**
 * 
 */
package de.tudresden.annotator.utils.automations;

import org.eclipse.swt.ole.win32.OleAutomation;
import org.eclipse.swt.ole.win32.Variant;

/**
 * @author Elvis Koci
 */
public class CollectionsUtils {
	
	/**
	 * Get the item having the specified name from a OleAutomation object. The latter is a collection of OLE Objects. 
	 * This method will fail if the OleAutomation does not have the "Item" property.
	 * @param automation
	 * @param itemName a string that represents the name of the item.
	 * @return
	 */
	public static OleAutomation getItemByName(OleAutomation automation, String itemName){
		
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
	 * @param index an integer that represents the index number of the item in the collection. 
	 * @return
	 */
	public static OleAutomation getItemByIndex(OleAutomation automation, int index){
		
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
