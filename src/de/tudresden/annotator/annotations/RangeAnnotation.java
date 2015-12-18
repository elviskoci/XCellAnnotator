/**
 * 
 */
package de.tudresden.annotator.annotations;

/**
 * @author Elvis Koci
 */
public class RangeAnnotation extends Annotation {

	private AnnotationClass annotationClass; 	
	private String annotationName;
	private String rangeAddress;
	private String sheetName;
	private int sheetIndex;
	private Annotation parent;
	
	/**
	 * @param annotationClass
	 * @param annotationName
	 * @param rangeAddress
	 * @param sheetName
	 * @param sheetIndex
	 */
	public RangeAnnotation(AnnotationClass annotationClass, String annotationName, String rangeAddress,
			String sheetName, int sheetIndex) {
		super();
		this.annotationClass = annotationClass;
		this.annotationName = annotationName;
		this.rangeAddress = rangeAddress;
		this.sheetName = sheetName;
		this.sheetIndex = sheetIndex;
	}


	/**
	 * @return the annotationClass
	 */
	public AnnotationClass getAnnotationClass() {
		return annotationClass;
	}

	/**
	 * @param annotationClass the annotationClass to set
	 */
	public void setAnnotationClass(AnnotationClass annotationClass) {
		this.annotationClass = annotationClass;
	}

	/**
	 * @return the annotationName
	 */
	public String getAnnotationName() {
		return annotationName;
	}

	/**
	 * @param annotationName the annotationName to set
	 */
	public void setAnnotationName(String annotationName) {
		this.annotationName = annotationName;
	}

	/**
	 * @return the rangeAddress
	 */
	public String getRangeAddress() {
		return rangeAddress;
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
	 * @param rangeAddress the rangeAddress to set
	 */
	public void setRangeAddress(String rangeAddress) {
		this.rangeAddress = rangeAddress;
	}
	
	/**
	 * @return the parent
	 */
	public Annotation getParent() {
		return parent;
	}


	/**
	 * @param parent the parent to set
	 */
	public void setParent(Annotation parent) {
		this.parent = parent;
	}


	@Override
	public boolean equals(Annotation obj) {
		
		if (obj instanceof RangeAnnotation) {	
			RangeAnnotation rangeAnnotation = (RangeAnnotation) obj;		
            return rangeAnnotation.getAnnotationName().compareToIgnoreCase(this.getAnnotationName())==0;
        }
        return false;
	}
	
	@Override
	public int hashCode() {
		return this.getAnnotationName().hashCode();
	}
}
