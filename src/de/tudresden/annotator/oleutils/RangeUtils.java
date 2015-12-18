/**
 * 
 */
package de.tudresden.annotator.oleutils;

import java.util.Arrays;

import org.eclipse.swt.ole.win32.OleAutomation;
import org.eclipse.swt.ole.win32.Variant;

/**
 * @author Elvis Koci
 */
public class RangeUtils {
	
	/**
	 * Get the address of the range
	 * @param rangeAutomation an OleAutomation to access a Range of cells
	 * @return a string that represents the address of the given range. 
	 */
	public static String getRangeAddress(OleAutomation rangeAutomation){
		
		int[] addressIds = rangeAutomation.getIDsOfNames(new String[]{"Address"}); 
		Variant addressVariant = rangeAutomation.getProperty(addressIds[0]);	
		String address = addressVariant.getString();
		addressVariant.dispose();
		
		return address;
	}
	
	
	/**
	 * Get areas from a range that represents a multi area selection  
	 * @param rangeAutomation an OleAutomation to access a Range of cells
	 * @return an OleAutomation that provides access to the collection of areas. 
	 */
	public static OleAutomation getAreas(OleAutomation rangeAutomation){
		
		int[] areasPropertyIds = rangeAutomation.getIDsOfNames(new String[]{"Areas"}); 
		Variant areasVariant = rangeAutomation.getProperty(areasPropertyIds[0]);	
		OleAutomation areasAutomation = areasVariant.getAutomation();
		areasVariant.dispose();
		
		return areasAutomation;
	}
	
	
	/**
	 * Get range value 
	 * @param rangeAutomation an OleAutomation to access a Range of cells
	 * @return a string that represents the value of the range
	 */
	public static String getValue(OleAutomation rangeAutomation){
		int[] valuePropertyIds = rangeAutomation.getIDsOfNames(new String[]{"Value"});
		
		Variant valueVariant =  rangeAutomation.getProperty(valuePropertyIds[0]);
		String value =  valueVariant.getString();
		valueVariant.dispose();
		
		return value;
	}
	
	
	/**
	 * Set a value to the range 
	 * @param rangeAutomation an OleAutomation to access a Range of cells
	 * @param value the string to set as value 
	 * @return true if the operation was successful, false otherwise
	 */
	public static boolean setValue(OleAutomation rangeAutomation, String value){
		
		int[] valuePropertyIds = rangeAutomation.getIDsOfNames(new String[]{"Value"});
		
		Variant valueVariant = new Variant(value);
		boolean isSuccess = rangeAutomation.setProperty(valuePropertyIds[0], valueVariant);
		valueVariant.dispose();
		
		return isSuccess;
	}
	
	
	/**
	 * Get the number of the first column in the first area in the specified range
	 * @param rangeAutomation an OleAutomation to access a Range of cells
	 * @return the number of the first column in the first area in the specified range
	 */
	public static String getRangeColumn(OleAutomation rangeAutomation){
		
		int[] columnPropertyIds = rangeAutomation.getIDsOfNames(new String[]{"Column"}); 
		Variant columnPropertyVariant = rangeAutomation.getProperty(columnPropertyIds[0]);	
		String column = columnPropertyVariant.getString();
		columnPropertyVariant.dispose();
		
		return column;
	}
	
	
	/**
	 * Get collection of columns in the range 
	 * @param rangeAutomation an OleAutomation to access a Range of cells
	 * @return an OleAutomation that provides access to the collection of columns in the range 
	 */
	public static OleAutomation getRangeColumns(OleAutomation rangeAutomation){
		
		int[] columnsPropertyIds = rangeAutomation.getIDsOfNames(new String[]{"Columns"}); 
		Variant columnsPropertyVariant = rangeAutomation.getProperty(columnsPropertyIds[0]);	
		OleAutomation columnsAutomation =  columnsPropertyVariant.getAutomation();
		columnsPropertyVariant.dispose();
		
		return columnsAutomation;
	}
	
	
	/**
	 * Get the number of the first row in the first area in the specified range
	 * @param rangeAutomation an OleAutomation to access a Range of cells
	 * @return the number of the first row in the first area in the specified range
	 */
	public static String getRangeRow(OleAutomation rangeAutomation){
		
		int[] rowPropertyIds = rangeAutomation.getIDsOfNames(new String[]{"Row"}); 
		Variant rowPropertyVariant = rangeAutomation.getProperty(rowPropertyIds[0]);	
		String row = rowPropertyVariant.getString();
		rowPropertyVariant.dispose();
		
		return row;
	}
	
	/**
	 * Get collection of rows in the range 
	 * @param rangeAutomation an OleAutomation to access a Range of cells
	 * @return an OleAutomation that provides access to the collection of rows in the range 
	 */
	public static OleAutomation getRangeRows(OleAutomation rangeAutomation){
		
		int[] rowsPropertyIds = rangeAutomation.getIDsOfNames(new String[]{"Rows"}); 
		Variant rowsPropertyVariant = rangeAutomation.getProperty(rowsPropertyIds[0]);	
		OleAutomation rowsAutomation = rowsPropertyVariant.getAutomation();
		rowsPropertyVariant.dispose();
		
		return rowsAutomation;
	}
	
	
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
	 * Filter range based on a criteria applied to a single field  
	 * @param rangeAutomation an OleAutomation to access a Range of cells
	 * @return true if the operation succeeded, false otherwise
	 */
	public static boolean filterRange(OleAutomation rangeAutomation, int field, String criteria1){
		
		int[] autoFilterMethodIds = rangeAutomation.getIDsOfNames(new String[]{"AutoFilter", "Field", "Criteria1"});
		
		Variant[] args = new Variant[2];
		args[0] = new Variant(field);
		args[1] = new Variant(criteria1);
		int argsIds[] = Arrays.copyOfRange(autoFilterMethodIds, 1, autoFilterMethodIds.length);
		
		Variant result = rangeAutomation.invoke(autoFilterMethodIds[0], args, argsIds);
		for (Variant arg : args) {
			arg.dispose();
		}
		
		if(result == null){
			return false;
		}
		
		result.dispose();
		return true;
	}
	
	
	/**
	 * Get special cells from the given range
	 * @param rangeAutomation an OleAutomation to access a Range of cells
	 * @param type an integer that represents the type of special cells to get. For more info see XlCellType enumeration. 
	 * @return an OleAutomation to access the range of special cells
	 */
	public static OleAutomation getSpecialCells(OleAutomation rangeAutomation, int type){
		
		int[] specialCellsMethodIds = rangeAutomation.getIDsOfNames(new String[]{"SpecialCells", "Type"});
		
		Variant[] args = new Variant[1];
		args[0] = new Variant(type);
		int argsIds[] = Arrays.copyOfRange(specialCellsMethodIds, 1, specialCellsMethodIds.length);
		
		Variant result = rangeAutomation.invoke(specialCellsMethodIds[0], args, argsIds);
		for (Variant arg : args)
			arg.dispose();
		
		if(result==null)
			return null; 
		
		OleAutomation specialCells = result.getAutomation();
		result.dispose();
		
		return specialCells;
	}
	
	
	/**
	 * Draw a border around the range
	 * @param rangeAutomation an OleAutomation to access a Range of cells
	 * @param lineStyle one of the constants of XlLineStyle 
	 * @param weight one of the constants of XlBorderWeight
	 * @param colorIndex the border color, as an index into the current color palette or as an XlColorIndex constant.
	 * @return true if operation succeeded, false otherwise
	 */
	public static boolean  drawBorderAroundRange(OleAutomation rangeAutomation, int lineStyle, double weight, long color){
		
		int[] borderAroundMethodIds = rangeAutomation.getIDsOfNames(new String[]{"BorderAround","LineStyle", "Weight", "Color"}); // "ColorIndex" 
		Variant methodParams[] = new Variant[3];
		methodParams[0] = new Variant(lineStyle); // line style (e.g., continuous, dashed ) 
		methodParams[1] = new Variant(weight); // border weight  (e.g., thick, thin )
		methodParams[2] = new Variant(color); // RGB color as a long value 
	
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
	 * Remove border around the range
	 * @param rangeAutomation an OleAutomation to access a Range of cells
	 * @return true if operation succeeded, false otherwise
	 */
	public static boolean removeBorderAroundRange(OleAutomation rangeAutomation){
		 
		int[] borderAroundMethodIds = rangeAutomation.getIDsOfNames(new String[]{"BorderAround", "LineStyle"});
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
	
	
	/**
	 * Delete the given range of cells 
	 * @param rangeAutomation an OleAutomation to access a Range of cells
	 * @return true if the operation succeeded, false otherwise
	 */
	public static boolean deleteRange(OleAutomation rangeAutomation){
		
		int[] deleteMethodIds = rangeAutomation.getIDsOfNames(new String[]{"Delete"});		
		Variant result = rangeAutomation.invoke(deleteMethodIds[0]);
		
		if(result == null){
			return false;
		}
		
		result.dispose();
		return true;
	}
}
