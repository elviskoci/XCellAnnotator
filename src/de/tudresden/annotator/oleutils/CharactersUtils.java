/**
 * 
 */
package de.tudresden.annotator.oleutils;

import org.eclipse.swt.ole.win32.OleAutomation;
import org.eclipse.swt.ole.win32.Variant;

/**
 * @author Elvis Koci
 */
public class CharactersUtils {
	
	/**
	 * Set Text using a Characters OleAutomation
	 * @param charactersAutomation an OleAutomation of a Characters object, which represents a range of characters within a shape’s text frame
	 * @param text the string to set as text
	 * @return true if operation was successful, false otherwise
	 */
	public static boolean setText(OleAutomation charactersAutomation, String text){		
		int[] textPropertyIds = charactersAutomation.getIDsOfNames(new String[]{"Text"}); 
		return charactersAutomation.setProperty(textPropertyIds[0], new Variant(text));
	}
	
	/**
	 * Get Font OleObject. This object can be used to change the font attributes of the text
	 * @param charactersAutomation an OleAutomation of a Characters object, which represents a range of characters within a shape’s text frame
	 * @return an OleAutomation that provides access to the Font Ole object. 
	 */
	public static OleAutomation getFontAutomation(OleAutomation charactersAutomation){
		
		int[] fontPropertyIds = charactersAutomation.getIDsOfNames(new String[]{"Font"});
		Variant fontVariant = charactersAutomation.getProperty(fontPropertyIds[0]);
		OleAutomation fontAutomation = fontVariant.getAutomation();
		fontVariant.dispose();
		
		return fontAutomation;
	}
}
