/**
 * 
 */
package de.tudresden.annotator.annotations;

/**
 * @author Elvis Koci
 */
public class WorksheetAnnotation extends Annotation {

	private String sheetName;
	private int sheetIndex;

	/**
	 * @param workbookName
	 * @param sheetName
	 * @param sheetIndex
	 */
	public WorksheetAnnotation(String sheetName, int sheetIndex) {
		this.sheetIndex = sheetIndex;
		this.sheetName = sheetName;
		
		String annotationId = sheetName+"_"+sheetIndex;
		this.setAnnotationId(annotationId);
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

	
	@Override
	public boolean equals(Annotation obj) {
		if (obj instanceof WorksheetAnnotation) {

			WorksheetAnnotation sheetAnnotation = (WorksheetAnnotation) obj;

			return sheetAnnotation.getSheetName().compareTo(this.getSheetName()) == 0
					&& sheetAnnotation.getSheetIndex() == this.getSheetIndex();
		}
		return false;
	}

	
	@Override
	public int hashCode() {
		return this.getSheetName().hashCode() + this.getSheetIndex();
	}

}
