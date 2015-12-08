/**
 * 
 */
package de.tudresden.annotator.utils.automations;

import org.eclipse.swt.ole.win32.OleAutomation;
import org.eclipse.swt.ole.win32.Variant;

/**
 * @author Elvis Koci
 */
public class TextFrameUtils {
	
	/**
	 * Set text horizontal alignment in the shape TextFrame  
	 * @param textFrameAutomation an OleAutomation that provides access to a TextFrame Ole object. It represents a textframe in shape object.
	 * @param alignment one of the XlVAlign constants
	 * @return true if operation succeeded, false otherwise
	 */
	public static boolean setVerticalAlignment(OleAutomation textFrameAutomation, int alignment){
		
		int[] verticalAlignmentPropertyIds = textFrameAutomation.getIDsOfNames(new String[]{"VerticalAlignment"});
		return textFrameAutomation.setProperty(verticalAlignmentPropertyIds[0], new Variant(alignment));
	}
	
	/**
	 * Set text vertical alignment in the shape TextFrame
	 * @param textFrameAutomation an OleAutomation that provides access to a TextFrame Ole object. It represents a textframe in shape object.
	 * @param alignment one of the XlHAlign constants
	 * @return true if operation succeeded, false otherwise
	 */
	public static boolean setHorizontalAlignment(OleAutomation textFrameAutomation, int alignment){
		
		int[] horizontalAlignmentPropertyIds = textFrameAutomation.getIDsOfNames(new String[]{"HorizontalAlignment"});
		return textFrameAutomation.setProperty(horizontalAlignmentPropertyIds[0], new Variant(alignment));
	}

	/**
	 * Get Characters OleAutomation. Use this object to set and format text in a TextFrame 
	 * @param textFrameAutomation an OleAutomation that provides access to a TextFrame Ole object. It represents a textframe in shape object.
	 * @return an OleAutomation that provides access to a Characters object.
	 */
	public static OleAutomation getCharactersAutomation(OleAutomation textFrameAutomation){
		
		int[] charactersPropertyIds = textFrameAutomation.getIDsOfNames(new String[]{"Characters"});
		Variant charactersVariant = textFrameAutomation.invoke(charactersPropertyIds[0]);
		OleAutomation charactersAutomation = charactersVariant.getAutomation();
		charactersVariant.dispose();
		
		return charactersAutomation;
	}

}
