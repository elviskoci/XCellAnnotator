/**
 * 
 */
package de.tudresden.annotator.oleutils;

import org.eclipse.swt.ole.win32.OleAutomation;
import org.eclipse.swt.ole.win32.Variant;

/**
 * @author Elvis Koci
 */
public class WorksheetFunctionUtils {
	
	/**
	 * Call an excel worksheet function (Ex. count, sum, min, max)
	 * @param worksheetFunction an OleAutomation that provides access to the worksheet functions
	 * @param name a string that represents the name of the function to call 
	 * @param args an array of variants that represent the arguments of the function to call 
	 * @return the result of the function
	 */
	public static Variant callFunction(OleAutomation worksheetFunction, String name, Variant[] args){
		
		int[] methodIds = worksheetFunction.getIDsOfNames(new String[]{name});
		if(methodIds==null){
			System.out.println("Method \""+name+"\" is not found for \"WorksheetFunction\" OLE Object!");
			return null;
		}
		
		Variant result = worksheetFunction.invoke(methodIds[0], args);
		 
		for (Variant var : args) {
			var.dispose();
		}
		
		// invoke the unprotect method  
		return result;
	}

}
