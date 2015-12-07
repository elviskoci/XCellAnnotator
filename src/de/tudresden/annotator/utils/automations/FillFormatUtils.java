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
	 * Set the color of the shape fill 
	 * @param fillFormatAutomation an OleAutomation that provides access to the FillFormat OLE Object
	 * @param color a long that represents a RGB color. Is calculated as B * 65536 + G * 256 + R
	 * @return true if operation succeeded, false otherwise
	 */
	public static boolean setShapeBackgroundColor(OleAutomation fillFormatAutomation, long color){
	
		int[] foreColorPropertyIds = fillFormatAutomation.getIDsOfNames(new String[]{"ForeColor"}); // BackColor does not work
		Variant foreColorVariant = fillFormatAutomation.getProperty(foreColorPropertyIds[0]);
		OleAutomation foreColorAutomation = foreColorVariant.getAutomation();

		int[] rgbPropertyIds = foreColorAutomation.getIDsOfNames(new String[]{"RGB"}); //alternative "SchemeColor" 
		foreColorAutomation.setProperty(rgbPropertyIds[0], new Variant(color)); 
	
		boolean isSuccess = fillFormatAutomation.setProperty(foreColorPropertyIds[0], foreColorVariant);			
		foreColorVariant.dispose();
		foreColorAutomation.dispose();
		
		return isSuccess;
	}
	
	
	/**
	 * Set the transparency of the shape fill
	 * @param fillFormatAutomation an OleAutomation that provides access to the FillFormat OLE Object
	 * @param transparency as a double. The expected values are between 0 and 1.
	 * @return true if operation succeeded, false otherwise
	 */
	public static boolean setShapeFillTransparency(OleAutomation fillFormatAutomation, double transparency){
		int[] transparencyPropertyIds = fillFormatAutomation.getIDsOfNames(new String[]{"Transparency"}); 
		return fillFormatAutomation.setProperty(transparencyPropertyIds[0], new Variant(transparency));	
	}
}
