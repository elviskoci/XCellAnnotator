/**
 * 
 */
package de.tudresden.annotator.annotations;

import java.util.Collection;
import java.util.HashMap;

/**
 * @author Elvis Koci
 */
public class WorkbookAnnotation {

	private String workbookName;
	private HashMap<String, WorksheetAnnotation> worksheetAnnotations;
	
	
	/**
	 * @param workbookName
	 */
	public WorkbookAnnotation() {
		this.worksheetAnnotations = new  HashMap<String, WorksheetAnnotation>();
	}
	
	
	/**
	 * @param workbookName
	 */
	public WorkbookAnnotation(String workbookName) {
		this.workbookName = workbookName;
		this.worksheetAnnotations = new  HashMap<String, WorksheetAnnotation>();
	}

	/**
	 * @param workbookName
	 * @param worksheetAnnotations
	 */
	public WorkbookAnnotation(String workbookName, HashMap<String, WorksheetAnnotation> worksheetAnnotations) {
		this.workbookName = workbookName;
		this.worksheetAnnotations = worksheetAnnotations;
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
	 * @return the worksheetAnnotations
	 */
	public HashMap<String, WorksheetAnnotation> getWorksheetAnnotations() {
		return worksheetAnnotations;
	}
	
	/**
	 * @param worksheetAnnotations the worksheetAnnotations to set
	 */
	public void setWorksheetAnnotations(HashMap<String, WorksheetAnnotation> worksheetAnnotations) {
		this.worksheetAnnotations = worksheetAnnotations;
	}
	
	
	/**
	 * 
	 * @param annotation
	 */
	public void addAnnotation(RangeAnnotation annotation){
		
		WorksheetAnnotation sheetAnnotation= this.worksheetAnnotations.get(annotation.getSheetName());
		
		if(sheetAnnotation==null){
			sheetAnnotation = new WorksheetAnnotation(annotation.getSheetName(), annotation.getSheetIndex());
			this.worksheetAnnotations.put(sheetAnnotation.getAnnotationId(), sheetAnnotation);
		}
		
		sheetAnnotation.addAnnotation(annotation.getAnnotationId(), annotation);
		sheetAnnotation.addAnnotationToSet(annotation.getAnnotationClass().getLabel(), annotation.getAnnotationId(), annotation);
	}
	
	
	/**
	 * 
	 * @param annotation
	 */
	public void removeAnnotation(RangeAnnotation annotation){
		WorksheetAnnotation sheetAnnotation= this.worksheetAnnotations.get(annotation.getSheetName());
		
		if(sheetAnnotation==null)
			return;
		
		sheetAnnotation.removeAnnotation(annotation.getAnnotationId());
		String classLabel = annotation.getAnnotationClass().getLabel();
		
		sheetAnnotation.removeAnnotationFromSet(classLabel, annotation.getAnnotationId());
	}
	
	
	/**
	 * 
	 * @param sheetName
	 * @param classLabel
	 * @param annotationId
	 */
	public void removeAnnotation(String sheetName, String classLabel, String annotationId){
		WorksheetAnnotation sheetAnnotation= this.worksheetAnnotations.get(sheetName);
		
		if(sheetAnnotation==null)
			return;
		
		sheetAnnotation.removeAnnotation(annotationId);
		sheetAnnotation.removeAnnotationFromSet(classLabel, annotationId);
	}
	
	
	/**
	 * 
	 * @param sheetName
	 * @param annotationId
	 * @return
	 */
	public Annotation getAnnotation(String sheetName, String annotationId){	
		WorksheetAnnotation sheetAnnotation= this.worksheetAnnotations.get(sheetName);
		
		if(sheetAnnotation==null)
			return null;
		
		return sheetAnnotation.getAnnotation(annotationId);
	}
	
	
	/**
	 * 
	 * @param sheetName
	 * @param classLabel
	 * @return
	 */
	public Collection<Annotation> getAnnotationsForClass(String sheetName, String classLabel){
		
		WorksheetAnnotation sheetAnnotation= this.worksheetAnnotations.get(sheetName);
		
		if(sheetAnnotation==null)
			return null;
		
		return sheetAnnotation.getAnnotationsByClass(classLabel);		
	}
}
