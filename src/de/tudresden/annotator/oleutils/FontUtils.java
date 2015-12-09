/**
 * 
 */
package de.tudresden.annotator.oleutils;

import org.eclipse.swt.ole.win32.OleAutomation;
import org.eclipse.swt.ole.win32.Variant;

/**
 * @author Elvis Koci
 */
public class FontUtils {
	
	/**
	 * Set text font size
	 * @param fontAutomation an OleAutomation that provides access to a Font OLE Object
	 * @param size the font size
	 * @return true if operation was successful, false otherwise
	 */
	public static boolean setFontSize(OleAutomation fontAutomation, int size){		
		int[] sizePropertyIds = fontAutomation.getIDsOfNames(new String[]{"Size"}); 
		return fontAutomation.setProperty(sizePropertyIds[0], new Variant(size));
	}
	
	/**
	 * Make text bold 
	 * @param fontAutomation an OleAutomation that provides access to a Font OLE Object
	 * @param bold true to make text bold, false otherwise 
	 * @return true if operation was successful, false if an error was encountered 
	 */
	public static boolean setBoldFont(OleAutomation fontAutomation, boolean bold){		
		int[] boldPropertyIds = fontAutomation.getIDsOfNames(new String[]{"Bold"}); 
		return fontAutomation.setProperty(boldPropertyIds[0], new Variant(bold));
	}
	
	/**
	 * Set font color for the text
	 * @param fontAutomation an OleAutomation that provides access to a Font OLE Object
	 * @param color a long that represents a RGB color. Is calculated as B * 65536 + G * 256 + R
	 * @return true if operation was successful, false otherwise
	 */
	public static boolean setFontColor(OleAutomation fontAutomation, long color){		
		int[] longPropertyIds = fontAutomation.getIDsOfNames(new String[]{"Color"}); 
		return fontAutomation.setProperty(longPropertyIds[0], new Variant(color));
	}

}
