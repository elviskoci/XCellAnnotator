/**
 * 
 */
package de.tudresden.annotator.utils.automations;

import java.util.Arrays;

import org.eclipse.swt.ole.win32.OleAutomation;
import org.eclipse.swt.ole.win32.Variant;

/**
 * @author Elvis Koci
 */
public class RangeUtils {
	
	
	/**
	 * Get the distance, in points, from the left edge of column A to the left edge of the range.
	 * @param rangeAutomation an OleAutomation to access a Range of cells
	 * @return the distance, in points, from the left edge of column A to the left edge of the range
	 */
	public static double getRangeLeftPosition(OleAutomation rangeAutomation){

		int[] leftPropertyIds = rangeAutomation.getIDsOfNames(new String[]{"Left"});
		Variant leftVariant=rangeAutomation.getProperty(leftPropertyIds[0]);
		double left = leftVariant.getDouble();
		leftVariant.dispose();
		return left;
	}
	
	
	/**
	 * Get the distance, in points, from the top edge of row 1 to the top edge of the range
	 * @param rangeAutomation an OleAutomation to access a Range of cells
	 * @return the distance, in points, from the top edge of row 1 to the top edge of the range
	 */
	public static double getRangeTopPosition(OleAutomation rangeAutomation){
		
		int[] topPropertyIds = rangeAutomation.getIDsOfNames(new String[]{"Top"});
		Variant topVariant=rangeAutomation.getProperty(topPropertyIds[0]);
		double top = topVariant.getDouble();
		topVariant.dispose();
		
		return top;
	}
	
	
	/**
	 * Get the height, in units, of the range.
	 * @param rangeAutomation an OleAutomation to access a Range of cells
	 * @return the height, in units, of the range.
	 */
	public static double getRangeHeight(OleAutomation rangeAutomation){
		
		int[] heightPropertyIds = rangeAutomation.getIDsOfNames(new String[]{"Height"});
		Variant heightVariant=rangeAutomation.getProperty(heightPropertyIds[0]);
		double height = heightVariant.getDouble();
		heightVariant.dispose();
		
		return height;
	}
	
	
	/**
	 * Get the width, in units, of the range.
	 * @param rangeAutomation an OleAutomation to access a Range of cells
	 * @return the width, in units, of the range.
	 */
	public static double getRangeWidth(OleAutomation rangeAutomation){
		
		int[] widthPropertyIds = rangeAutomation.getIDsOfNames(new String[]{"Width"});
		Variant widthVariant=rangeAutomation.getProperty(widthPropertyIds[0]);
		double width = widthVariant.getDouble();
		widthVariant.dispose();
		
		return width;
	}
	
	
	/**
	 * Draw a border around the range
	 * @param rangeAutomation an OleAutomation to access a Range of cells
	 * @param lineStyle one of the constants of XlLineStyle 
	 * @param weight one of the constants of XlBorderWeight
	 * @param colorIndex the border color, as an index into the current color palette or as an XlColorIndex constant.
	 * @return true if operation succeeded, false otherwise
	 */
	public static boolean  drawBorderAroundRange(OleAutomation rangeAutomation, int lineStyle, int weight, int colorIndex){
		
		int[] borderAroundMethodIds = rangeAutomation.getIDsOfNames(new String[]{"BorderAround","LineStyle", "Weight", "ColorIndex"}); // "Color"
		Variant methodParams[] = new Variant[3];
		methodParams[0] = new Variant(lineStyle); // line style (e.g., continuous, dashed ) 
		methodParams[1] = new Variant(weight); // border weight  (e.g., thick, thin )
		methodParams[2] = new Variant(colorIndex); // index into the current color palette
	
		int[] paramIds = Arrays.copyOfRange(borderAroundMethodIds, 1, borderAroundMethodIds.length);
		Variant result = rangeAutomation.invoke(borderAroundMethodIds[0], methodParams, paramIds);
		
		for (Variant v : methodParams) {
			v.dispose();
		}	
		
		if(result==null){
			return false;
		}
		
		result.dispose();
		return true;
	}
	
	/**
	 * Erase border around the range
	 * @param rangeAutomation an OleAutomation to access a Range of cells
	 * @return true if operation succeeded, false otherwise
	 */
	public static boolean eraseBorderAroundRange(OleAutomation rangeAutomation){
		 
		int[] borderAroundMethodIds = rangeAutomation.getIDsOfNames(new String[]{"BorderAround","LineStyle"});
		Variant methodParams[] = new Variant[1];
		
		int xlLineStyleNone = -4142; // no line
		methodParams[0] = new Variant(xlLineStyleNone);  
	
		int[] paramIds = Arrays.copyOfRange(borderAroundMethodIds, 1, borderAroundMethodIds.length);
		Variant result = rangeAutomation.invoke(borderAroundMethodIds[0], methodParams, paramIds);
		
		for (Variant v : methodParams) {
			v.dispose();
		}	
		
		if(result==null){
			return false;
		}
		
		result.dispose();
		return true;
	}
}
