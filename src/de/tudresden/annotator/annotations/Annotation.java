/**
 * 
 */
package de.tudresden.annotator.annotations;

import java.util.Collection;
import java.util.HashMap;

/**
 * @author Elvis Koci
 */
public abstract class Annotation {
	
	private String annotationId; 
	private HashMap <String, HashMap<String, Annotation>> annotationsByClass ; 
	private HashMap <String, Annotation> allAnnotations;
	
		
	public Annotation(){
		this.annotationsByClass =  new HashMap<String, HashMap<String, Annotation>>();
		this.allAnnotations =  new HashMap<String, Annotation>();
	}
	
	/**
	 * Add a new annotation
	 * @param key a string that is used as an id (key) for the annotation object
	 * @param annotation the annotation object to add
	 */
	public void addAnnotation(String key, Annotation annotation) {
		this.allAnnotations.put(key, annotation);
	}

	/**
	 * Remove annotation
	 * @param key a string that is used as an id (key) for the annotation object
	 */
	public void removeAnnotation(String key) {
		this.allAnnotations.remove(key);
	}
	
	/**
	 * Get the annotation by id
	 * @param key a string that is used as an id (key) for the annotation object
	 * @return
	 */
	public Annotation getAnnotation(String key){
		return this.allAnnotations.get(key);
	}

	
	/**
	 * Get all the annotations that have the given class label
	 * @param label a string that represents the label of a Annotation class
	 * @return a set of annotation objects
	 */
	public Collection<Annotation> getAnnotationsByClass(String label){
		HashMap<String, Annotation> map = annotationsByClass.get(label);
		if(map==null)
			return null;
		
		return map.values();
	}
	
	/**
	 * Add an annotation to the set containing annotations of the same class 
	 * @param label a string that represents the label of a Annotation class
	 * @param annotation the object that represents the annotation to add
	 */
	public void addAnnotationToSet(String label, String annotationId, Annotation annotation){
				
		HashMap<String, Annotation>  map = annotationsByClass.get(label);
		if(map == null){
			map = new HashMap<String, Annotation>();
		}
		
		map.put(annotationId, annotation);
		
		this.annotationsByClass.put(label, map);
	}
	
	/**
	 * Remove an annotation from the set containing annotations of the same class 
	 * @param label a string that represents the label of a Annotation class
	 * @param annotationId a string that represents the id of the annotation
	 */
	public void removeAnnotationFromSet(String label, String annotationId){
		
		HashMap<String, Annotation>  map = annotationsByClass.get(label);
		if(map == null)
			return;
		
		map.remove(annotationId);
	}

		
	/**
	 * @return the annotationId
	 */
	public String getAnnotationId() {
		return annotationId;
	}

	/**
	 * @param annotationId the annotationId to set
	 */
	protected void setAnnotationId(String annotationId) {
		this.annotationId = annotationId;
	}

	public abstract boolean equals(Annotation obj);

	public abstract int hashCode();
	
}
