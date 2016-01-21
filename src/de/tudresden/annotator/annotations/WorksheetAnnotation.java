/**
 * 
 */
package de.tudresden.annotator.annotations;

/**
 * @author Elvis Koci
 */
public class WorksheetAnnotation extends DependentAnnotation<WorkbookAnnotation> {

	private String workbookName;
	private String sheetName;
	private int sheetIndex;
	private boolean isCompleted = false;
	private boolean isNotApplicable = false;

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
		return isNotApplicable;
	}

	/**
	 * @param isNotApplicable the isNotApplicable to set
	 */
	public void setNotApplicable(boolean isNotApplicable) {
		this.isNotApplicable = isNotApplicable;
	}
	
	@Override 
	public String toString() {
		return this.getSheetName()+" = "+this.allAnnotations.values(); 
	}

	@Override
	public boolean equals(Annotation<RangeAnnotation> annotation) {
		
		if(!(annotation instanceof WorksheetAnnotation))
			return false;
		
		WorksheetAnnotation sa = (WorksheetAnnotation) annotation;
		
		if(sa.getSheetName().compareTo(this.sheetName)!=0)
			return false;
		
		if(!(sa.getAllAnnotations().equals(this.allAnnotations)))
			return false;
		
		if(sa.isCompleted()!=this.isCompleted)
			return false;
			
		if(sa.isNotApplicable()!=this.isNotApplicable)
			return false;
		
		return true;
	}

	@Override
	public int hashCode() {
		int hash = this.getSheetName().hashCode() + (this.isCompleted?1:0) + (this.isNotApplicable?1:0) ;
		
		for (RangeAnnotation val : this.allAnnotations.values()) {
			hash = hash + val.hashCode();
		}
		
		return hash;
	}
}
