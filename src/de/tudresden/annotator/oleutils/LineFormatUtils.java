/**
 * 
 */
package de.tudresden.annotator.oleutils;

import org.eclipse.swt.ole.win32.OleAutomation;
import org.eclipse.swt.ole.win32.Variant;

/**
 * @author Elvis Koci
 */
public class LineFormatUtils {
	
	/**
	 * Set the line style
	 * @param lineFormatAuto an OleAutomation object that provides access to the LineFormat object functionalities.
	 * @param style one of the constant values from MsoLineStyle enumeration 
	 * @return true if the operation was successful, false otherwise
	 */	
	public static boolean setLineStyle(OleAutomation lineFormatAuto, int style ){
		
		int stylePropertyIds[] = lineFormatAuto.getIDsOfNames(new String[] {"Style"});
		Variant styleVariant = new Variant(style); 
		boolean isSuccess = lineFormatAuto.setProperty(stylePropertyIds[0], styleVariant);
		styleVariant.dispose();
		
		return isSuccess;	
	}
	
	/**
	 * Set the weight of the line 
	 * @param lineFormatAuto an OleAutomation object that provides access to the LineFormat object functionalities.
	 * @param weight a double value that represents the weight of the line 
	 * @return true if the operation was successful, false otherwise
	 */	
	public static boolean setLineWeight(OleAutomation lineFormatAuto, double weight ){
		
		int weightPropertyIds[] = lineFormatAuto.getIDsOfNames(new String[] {"Weight"});
		Variant weightVariant = new Variant(weight); 
		boolean isSuccess = lineFormatAuto.setProperty(weightPropertyIds[0], weightVariant);
		weightVariant.dispose();
		
		return isSuccess;	
	}
	
	/**
	 * Set if line is visible or not
	 * @param lineFormatAuto an OleAutomation object that provides access to the LineFormat object functionalities.
	 * @param visible true if the line should be visible, false if the line should be hidden 
	 * @return true if the operation was successful, false otherwise
	 */	
	public static boolean setLineVisibility(OleAutomation lineFormatAuto, boolean visible ){
		
		int visiblePropertyIds[] = lineFormatAuto.getIDsOfNames(new String[] {"Visible"});
		Variant visibleVariant = new Variant(visible); 
		boolean isSuccess = lineFormatAuto.setProperty(visiblePropertyIds[0], visibleVariant);
		visibleVariant.dispose();
		
		return isSuccess;	
	}
	
	
	/**
	 * Set the transparency of the line 
	 * @param lineFormatAuto an OleAutomation object that provides access to the LineFormat object functionalities.
	 * @param transparency a double value that represents the line transparency
	 * @return true if the operation was successful, false otherwise
	 */	
	public static boolean setLineTransparency(OleAutomation lineFormatAuto, double transparency ){
		
		int transparencyPropertyIds[] = lineFormatAuto.getIDsOfNames(new String[] {"Transparency"});
		Variant transparencyVariant = new Variant(transparency); 
		boolean isSuccess = lineFormatAuto.setProperty(transparencyPropertyIds[0], transparencyVariant);
		transparencyVariant.dispose();
		
		return isSuccess;	
	}
}
