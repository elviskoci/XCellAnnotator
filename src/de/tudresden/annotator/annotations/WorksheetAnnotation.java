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
	boolean isCompleted = false;
	boolean notApplicable = false;

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

	/**
	 * @return the isCompleted
	 */
	public boolean isCompleted() {
		return isCompleted;
	}

	/**
	 * @param isCompleted the isCompleted to set
	 */
	public void setCompleted(boolean isCompleted) {
		this.isCompleted = isCompleted;
	}

	/**
	 * @return the isIrrelevant
	 */
	public boolean isNotApplicable() {
		return notApplicable;
	}

	/**
	 * @param isNotApplicable the isNotApplicable to set
	 */
	public void setNotApplicable(boolean isNotApplicable) {
		this.notApplicable = isNotApplicable;
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
