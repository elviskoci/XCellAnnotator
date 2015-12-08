/**
 * 
 */
package de.tudresden.annotator.utils.automations;

import org.eclipse.swt.ole.win32.OleAutomation;
import org.eclipse.swt.ole.win32.Variant;

/**
 * @author Elvis Koci
 */
public class FillFormatUtils {
	
	
	/**
	 * Show or hide the fill of the shape.
	 * @param fillFormatAutomation an OleAutomation that provides access to the FillFormat OLE Object
	 * @param visible true to show the fill, false to hide it
	 * @return true if operation succeeded, false otherwise
	 */
	public static boolean setVisible(OleAutomation fillFormatAutomation, boolean visible){
		
		int[] visiblePropertyIds = fillFormatAutomation.getIDsOfNames(new String[]{"Visible"}); 
		Variant visiblePropertyVariant = new Variant(visible); 
		boolean isSuccess = fillFormatAutomation.setProperty(visiblePropertyIds[0], visiblePropertyVariant);
		visiblePropertyVariant.dispose();
	
		return isSuccess;
	}
	
	
	/**
	 * Set the transparency of the shape fill
	 * @param fillFormatAutomation an OleAutomation that provides access to the FillFormat OLE Object
	 * @param transparency as a double. The expected values are between 0 and 1.
	 * @return true if operation succeeded, false otherwise
	 */
	public static boolean setFillTransparency(OleAutomation fillFormatAutomation, double transparency){
		int[] transparencyPropertyIds = fillFormatAutomation.getIDsOfNames(new String[]{"Transparency"}); 
		return fillFormatAutomation.setProperty(transparencyPropertyIds[0], new Variant(transparency));	
	}
}
