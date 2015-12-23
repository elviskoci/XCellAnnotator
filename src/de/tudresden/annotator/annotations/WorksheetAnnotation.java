/**
 * 
 */
package de.tudresden.annotator.annotations;

import org.json.JSONObject;

/**
 * @author Elvis Koci
 */
public class WorksheetAnnotation extends DependentAnnotation<WorkbookAnnotation> {

	private String workbookName;
	private String sheetName;
	private int sheetIndex;

	/**
	 * @param workbookName
	 * @param sheetName
	 * @param sheetIndex
	 */
	public WorksheetAnnotation(String workbookName, String sheetName, int sheetIndex) {
		this.workbookName = workbookName;
		this.sheetIndex = sheetIndex;
		this.sheetName = sheetName;
	}
		
	/**
	 * @param sheetName
	 * @param sheetIndex
	 */
	public WorksheetAnnotation(String sheetName, int sheetIndex) {
		this.sheetIndex = sheetIndex;
		this.sheetName = sheetName;
	}

	/**
	 * @return the workbookName
	 */
	public String getWorkbookName() {
		return workbookName;
	}

	/**
	 * @param workbookName the workbookName to set
	 */
	public void setWorkbookName(String workbookName) {
		this.workbookName = workbookName;
	}

	/**
	 * @return the sheetName
	 */
	public String getSheetName() {
		return sheetName;
	}

	/**
	 * @param sheetName the sheetName to set
	 */
	public void setSheetName(String sheetName) {
		this.sheetName = sheetName;
	}

	/**
	 * @return the sheetIndex
	 */
	public int getSheetIndex() {
		return sheetIndex;
	}

	/**
	 * @param sheetIndex the sheetIndex to set
	 */
	public void setSheetIndex(int sheetIndex) {
		this.sheetIndex = sheetIndex;
	}
	
	public static String generateKey(String sheetName, int sheetIndex) {
		//return sheetName+"_"+sheetIndex;
		return sheetName;
	}
	
	
	@Override 
	public String toString() {
		// JSONObject json = new JSONObject(this.allAnnotations);
		return this.allAnnotations.toString(); 
	}
	
//	@Override
//	public boolean equals(Annotation<WorkbookAnnotation, RangeAnnotation> annotation) {
//		
//		if (annotation instanceof WorksheetAnnotation) {
//			WorksheetAnnotation sheetAnnotation = (WorksheetAnnotation) annotation;
//
//			return sheetAnnotation.getSheetName().compareTo(this.getSheetName()) == 0
//					&& sheetAnnotation.getSheetIndex() == this.getSheetIndex();
//		}
//		return false;
//	}
//
//	
//	@Override
//	public int hashCode() {
//		return this.getSheetName().hashCode() + this.getSheetIndex();
//	}
}
