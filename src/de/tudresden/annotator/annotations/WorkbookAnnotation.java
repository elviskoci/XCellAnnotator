/**
 * 
 */
package de.tudresden.annotator.annotations;

import java.util.ArrayList;
import java.util.Collection;
import java.util.HashMap;

import de.tudresden.annotator.annotations.utils.AnnotationHandler;

/**
 * @author Elvis Koci
 */
public class WorkbookAnnotation extends Annotation<WorkbookAnnotation, RangeAnnotation>{

	private String workbookName;
	private HashMap<String, WorksheetAnnotation> worksheetAnnotations;
	
	/**
	 * @param workbookName
	 */
	public WorkbookAnnotation() {
		this.worksheetAnnotations = new  HashMap<String, WorksheetAnnotation>();
		this.setParent(null);
	}
	
	
	/**
	 * @param workbookName
	 */
	public WorkbookAnnotation(String workbookName) {
		this.workbookName = workbookName;
		this.worksheetAnnotations = new  HashMap<String, WorksheetAnnotation>();
		this.setParent(null);
	}

	/**
	 * @param workbookName
	 * @param worksheetAnnotations
	 */
	public WorkbookAnnotation(String workbookName, HashMap<String, WorksheetAnnotation> worksheetAnnotations) {
		this.workbookName = workbookName;
		this.worksheetAnnotations = worksheetAnnotations;
		this.setParent(null);
	}
	
	/**
	 * Add a new RangeAutomation
	 * @param sheetName the name of the worksheet where the RangeAnnotation is placed 
	 * @param sheetIndex the index of the worksheet where the RangeAnnotation is placed 
	 * @param annotationClass the AnnotationClass that this RangeAnnotation is member of
	 * @param name a string that represents the name of the RangeAnnotation
	 * @param rangeAddress the address of the range that was annotated 
	 */
	public void addRangeAnnotation(String sheetName, int sheetIndex, AnnotationClass annotationClass, String name, String rangeAddress ){
		
		String sheetKey = sheetName;
		WorksheetAnnotation sheetAnnotation= this.worksheetAnnotations.get(sheetKey);
		
		if(sheetAnnotation==null){
			sheetAnnotation = new WorksheetAnnotation(sheetName, sheetIndex);
			this.worksheetAnnotations.put(sheetKey, sheetAnnotation);
		}
		
		RangeAnnotation rangeAnnotation= new RangeAnnotation(sheetName, sheetIndex, annotationClass, name, rangeAddress);
		
		sheetAnnotation.addAnnotation(name, rangeAnnotation);
		sheetAnnotation.addAnnotationToBucket(annotationClass.getLabel(), name, rangeAnnotation);
		
		this.addAnnotation(name, rangeAnnotation);
		this.addAnnotationToBucket(annotationClass.getLabel(), name, rangeAnnotation);		
	}
	
	
	/**
	 * Add a RangeAnnotation 
	 * @param rangeAnnotation an object that represents a RangeAnnotation
	 */
	public void addRangeAnnotation(RangeAnnotation rangeAnnotation){
		
		String sheetKey = rangeAnnotation.getSheetName();
		WorksheetAnnotation sheetAnnotation= this.worksheetAnnotations.get(sheetKey);
		
		if(sheetAnnotation==null){
			sheetAnnotation = new WorksheetAnnotation(rangeAnnotation.getSheetName(), rangeAnnotation.getSheetIndex());
			this.worksheetAnnotations.put(sheetKey, sheetAnnotation);
		}
		
		sheetAnnotation.addAnnotation(rangeAnnotation.getName(), rangeAnnotation);
		sheetAnnotation.addAnnotationToBucket(rangeAnnotation.getAnnotationClass().getLabel(), rangeAnnotation.getName(), rangeAnnotation);
		
		this.addAnnotation(rangeAnnotation.getName(), rangeAnnotation);
		this.addAnnotationToBucket(rangeAnnotation.getAnnotationClass().getLabel(), rangeAnnotation.getName(), rangeAnnotation);
	}
	
	/**
	 * Get the RangeAutomation based on the worksheet key and annotation key
	 * @param sheetKey a string that represents the id (key) of the worksheet where the RangeAnnotation is placed 
	 * @param annotationKey a string that is used as key for the annotation object 
	 * @return the RangeAnnotation object that corresponds to the given arguments  
	 */
	public RangeAnnotation getRangeAnnotation(String sheetKey, String annotationKey){	
		WorksheetAnnotation sheetAnnotation= this.worksheetAnnotations.get(sheetKey);
		
		if(sheetAnnotation==null)
			return null;
		
		return sheetAnnotation.getAnnotation(annotationKey);
	}
	
	
	/**
	 * Get the collection of RangeAnnotations for the given Worksheet key and AnnotationClass label
	 * @param sheetKey a string that represents the id (key) of the worksheet where the RangeAnnotation is placed
	 * @param classLabel the label of the AnnotationClass that this RangeAnnotation is member of
	 * @return a collection of RangeAutomations that correspond to the given arguments
	 */
	public Collection<RangeAnnotation> getRangeAnnotationsFromBucket(String sheetKey, String classLabel){
		
		WorksheetAnnotation sheetAnnotation= this.worksheetAnnotations.get(sheetKey);
		
		if(sheetAnnotation==null)
			return null;
		
		return sheetAnnotation.getAnnotationsByClass(classLabel);		
	}
	
	
	/**
	 * Remove a RangeAutomation 
	 * @param rangeAnnotation an object that represents a RangeAnnotation
	 */
	public void removeRangeAnnotation(RangeAnnotation rangeAnnotation){
		String sheetKey = rangeAnnotation.getSheetName(); 
		WorksheetAnnotation sheetAnnotation= this.worksheetAnnotations.get(sheetKey);
		
		if(sheetAnnotation==null)
			return;
		
		String classLabel = rangeAnnotation.getAnnotationClass().getLabel();
		
		sheetAnnotation.removeAnnotation(rangeAnnotation.getName());
		sheetAnnotation.removeAnnotationFromBucket(classLabel, rangeAnnotation.getName());
		
		this.removeAnnotation(rangeAnnotation.getName());
		this.removeAnnotationFromBucket(classLabel, rangeAnnotation.getName());
	}
	
	
	/**
	 * Remove a RangeAutomation 
	 * @param sheetKey a string that represents the id (key) of the worksheet where the RangeAnnotation is placed
	 * @param classLabel the label of the AnnotationClass that this RangeAnnotation is member of
	 * @param annotationKey a string that is used as key for the annotation object 
	 */
	public void removeRangeAnnotation(String sheetKey, String classLabel, String annotationKey){
		WorksheetAnnotation sheetAnnotation= this.worksheetAnnotations.get(sheetKey);
		
		if(sheetAnnotation==null)
			return;
		
		sheetAnnotation.removeAnnotation(annotationKey);
		sheetAnnotation.removeAnnotationFromBucket(classLabel, annotationKey);
		
		this.removeAnnotation(annotationKey);
		this.removeAnnotationFromBucket(classLabel, annotationKey);
	}
		
	/**
	 * Remove all RangeAutomations belonging to the specified Workbook and AnnotationClass
	 * @param sheetKey a string that represents the id (key) of the worksheet where the RangeAnnotation is placed
	 * @param classLabel the label of the AnnotationClass that this RangeAnnotation is member of
	 * 
	 */
	public void emptyRangeAnnotationBucket(String sheetKey, String classLabel){
		
		WorksheetAnnotation sheetAnnotation= this.worksheetAnnotations.get(sheetKey);
		
		if(sheetAnnotation==null)
			return;
		
		ArrayList<RangeAnnotation>  annotations= new ArrayList<RangeAnnotation>(sheetAnnotation.getAnnotationsByClass(classLabel));	
		for (RangeAnnotation rangeAnnotation : annotations) {
			this.removeAnnotationFromBucket(classLabel, rangeAnnotation.getKey());
			this.removeAnnotation(rangeAnnotation.getKey());
		}
		
		sheetAnnotation.removeAllAnnotationsOfClass(classLabel);
	}
	
	
	/**
	 * Remove all RangeAutomation belonging to the worksheet with the given key
	 * @param sheetKey a string that represents the id (key) of the worksheet where the RangeAnnotation is placed
	 */
	public void removeAllAnnotationFromSheet(String sheetKey){
		WorksheetAnnotation sheetAnnotation= this.worksheetAnnotations.get(sheetKey);
		sheetAnnotation.removeAllAnnotations();
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

	
	/* (non-Javadoc)
	 * @see de.tudresden.annotator.annotations2.Annotation#getKey()
	 */
	@Override
	protected String getKey() {
		return this.workbookName;
	}

	
	
//	@Override
//	public boolean equals( Annotation<WorkbookAnnotation, RangeAnnotation> annotation) {
//		if (annotation instanceof WorkbookAnnotation) {	
//			WorkbookAnnotation workbookAnnotation = (WorkbookAnnotation) annotation;		
//            return workbookAnnotation.getWorkbookName().compareToIgnoreCase(this.getWorkbookName())==0;
//        }
//        return false;
//	}
//
//	@Override
//	public int hashCode() {
//		return this.getWorkbookName().hashCode();
//	}
	
}
