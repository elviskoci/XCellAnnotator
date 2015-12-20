/**
 * 
 */
package de.tudresden.annotator.annotations;

/**
 * @author Elvis Koci
 */
public class WorksheetAnnotation extends Annotation<WorkbookAnnotation, RangeAnnotation> {

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

	/* (non-Javadoc)
	 * @see de.tudresden.annotator.annotations2.Annotation#getKey()
	 */
	@Override
	protected String getKey() {
		return this.sheetName ;
	}
	
	protected static String getKey(String sheetName, String sheetIndex) {
		return sheetName+"_"+sheetIndex;
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
