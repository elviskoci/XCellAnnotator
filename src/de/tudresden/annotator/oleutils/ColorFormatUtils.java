/**
 * 
 */
package de.tudresden.annotator.oleutils;

import org.eclipse.swt.ole.win32.OleAutomation;
import org.eclipse.swt.ole.win32.Variant;

/**
 * @author Elvis Koci
 */
public class ColorFormatUtils {
	
	/**
	 * Set the fore color of the shape fill 
	 * @param automation an OleAutomation that has the "ForeColor" property
	 * @param color a long that represents a RGB color. Is calculated as B * 65536 + G * 256 + R
	 * @return true if operation succeeded, false otherwise
	 */
	public static boolean setForeColor(OleAutomation automation, long color){
	
		int[] foreColorPropertyIds = automation.getIDsOfNames(new String[]{"ForeColor"}); 
		
		if(foreColorPropertyIds==null)	{		
			System.out.println("The given OleObject does not have the \"ForeColor\" property");
			return false;
		}
		
		Variant foreColorVariant = automation.getProperty(foreColorPropertyIds[0]);
		OleAutomation foreColorAutomation = foreColorVariant.getAutomation();

		int[] rgbPropertyIds = foreColorAutomation.getIDsOfNames(new String[]{"RGB"}); //alternatively use "SchemeColor" 
		foreColorAutomation.setProperty(rgbPropertyIds[0], new Variant(color)); 
	
		boolean isSuccess = automation.setProperty(foreColorPropertyIds[0], foreColorVariant);			
		foreColorVariant.dispose();
		foreColorAutomation.dispose();
		
		return isSuccess;
	}
	
	
	/**
	 * Set the back color of the shape fill 
	 * @param automation an OleAutomation that has the "BackColor" property
	 * @param colorIndex a long that represents a RGB color. Is calculated as B * 65536 + G * 256 + R 
	 * @return true if operation succeeded, false otherwise
	 */
	public static boolean setBackColor(OleAutomation automation,  long color){
	
		int[] backColorPropertyIds = automation.getIDsOfNames(new String[]{"BackColor"}); 
		
		if(backColorPropertyIds==null)	{		
			System.out.println("The given OleObject does not have the \"BackColor\" property");
			return false;
		}
		
		Variant backColorVariant = automation.getProperty(backColorPropertyIds[0]);
		OleAutomation backColorAutomation = backColorVariant.getAutomation();
	
		int[] rgbPropertyIds = backColorAutomation.getIDsOfNames(new String[]{"RGB"}); //alternatively use "SchemeColor" 
		backColorAutomation.setProperty(rgbPropertyIds[0], new Variant(color)); 
	
		boolean isSuccess = automation.setProperty(backColorPropertyIds[0], backColorVariant);
		
		backColorVariant.dispose();
		backColorAutomation.dispose();
		
		return isSuccess;
	}
	
	
	/**
	 * Set the fore color of the shape fill 
	 * @param automation an OleAutomation that has the "ForeColor" property
	 * @param colorIndex an integer that represents the index of the color in the current color palette. 
	 * @return true if operation succeeded, false otherwise
	 */
	public static boolean setForeColor(OleAutomation automation, int colorIndex){
	
		int[] foreColorPropertyIds = automation.getIDsOfNames(new String[]{"ForeColor"});
		
		if(foreColorPropertyIds==null)	{		
			System.out.println("The given OleObject does not have the \"ForeColor\" property");
			return false;
		}
		
		Variant foreColorVariant = automation.getProperty(foreColorPropertyIds[0]);
		OleAutomation foreColorAutomation = foreColorVariant.getAutomation();
	
		int[] schemeColorPropertyIds = foreColorAutomation.getIDsOfNames(new String[]{"SchemeColor"}); //alternatively use "RGB" 
		foreColorAutomation.setProperty(schemeColorPropertyIds[0], new Variant(colorIndex)); 
	
		boolean isSuccess = automation.setProperty(foreColorPropertyIds[0], foreColorVariant);			
		foreColorVariant.dispose();
		foreColorAutomation.dispose();
		
		return isSuccess;
	}
	
	
	/**
	 * Set the back color of the shape fill 
	 * @param automation an OleAutomation that has the "BackColor" property
	 * @param colorIndex an integer that represents the index of the color in the current color palette. 
	 * @return true if operation succeeded, false otherwise
	 */
	public static boolean setBackColor(OleAutomation automation, int colorIndex){
	
		int[] backColorPropertyIds = automation.getIDsOfNames(new String[]{"BackColor"}); 
		
		if(backColorPropertyIds==null)	{		
			System.out.println("The given OleObject does not have the \"BackColor\" property");
			return false;
		}
		
		Variant backColorVariant = automation.getProperty(backColorPropertyIds[0]);
		OleAutomation backColorAutomation = backColorVariant.getAutomation();
	
		int[] schemeColorPropertyIds = backColorAutomation.getIDsOfNames(new String[]{"SchemeColor"}); //alternatively use "RGB" 
		backColorAutomation.setProperty(schemeColorPropertyIds[0], new Variant(colorIndex)); 
	
		boolean isSuccess = automation.setProperty(backColorPropertyIds[0], backColorVariant);
		
		backColorVariant.dispose();
		backColorAutomation.dispose();
		
		return isSuccess;
	}
	
	/**
	 * Get RGB color as long value
	 * @param red value of red color
	 * @param green value of green color
	 * @param blue value of blue color
	 * @return a long value that represents an RGB color 
	 */
	public static long getRGBColorAsLong(int red, int green, int blue){	
		return blue * 65536 + green * 256 + red;
	}
}
