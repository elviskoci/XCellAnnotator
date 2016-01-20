/**
 * 
 */
package de.tudresden.annotator.annotations;

import de.tudresden.annotator.annotations.utils.AnnotationHandler;

/**
 * @author Elvis Koci
 */
public class RangeAnnotation extends DependentAnnotation<DependentAnnotation<?>> {

	private String sheetName;
	private int sheetIndex;
	private AnnotationClass annotationClass; 	
	private String name;
	private String rangeAddress;
	
	
	/**
	 * Create a new RangeAnnotation
	 * @param sheetName the name of the worksheet where the RangeAnnotation is placed 
	 * @param sheetIndex the index of the worksheet where the RangeAnnotation is placed 
	 * @param annotationClass the AnnotationClass that this RangeAnnotation is member of
	 * @param name a string that represents the name of the RangeAnnotation
	 * @param rangeAddress the address of the range that was annotated 
	 */
	public RangeAnnotation(String sheetName, int sheetIndex, AnnotationClass annotationClass, String name, String rangeAddress ) {
		this.annotationClass = annotationClass;
		this.name = name;
		this.rangeAddress = rangeAddress;
		this.sheetName = sheetName;
		this.sheetIndex = sheetIndex;	
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
	 * @return the name
	 */
	public String getName() {
		return name;
	}

	/**
	 * @param name the name to set
	 */
	public void setName(String name) {
		this.name = name;
	}

	/**
	 * @return the rangeAddress
	 */
	public String getRangeAddress() {
		return rangeAddress;
	}

	/**
	 * @param rangeAddress the rangeAddress to set
	 */
	public void setRangeAddress(String rangeAddress) {
		this.rangeAddress = rangeAddress;
	}

		
	public static String generateKey(String sheetName, String classLabel, String rangeAddress) {
		// TODO: Remove this method
		return AnnotationHandler.getStartOfRangeAnnotationName(sheetName)+"_"+classLabel+"_"+rangeAddress;
	}
	
	@Override 
	public String toString() {
			return this.name+" = "+this.allAnnotations.values().toString();
	}

//	@Override
//	public boolean equals(Annotation <RangeAnnotation, RangeAnnotation> annotation) {
//		if (annotation instanceof RangeAnnotation) {	
//			RangeAnnotation rangeAnnotation = (RangeAnnotation) annotation;		
//            return rangeAnnotation.getName().compareToIgnoreCase(this.getName())==0;
//        }
//        return false;
//	}
//
//	
//	@Override
//	public int hashCode() {
//		return this.getName().hashCode();
//	}
}
