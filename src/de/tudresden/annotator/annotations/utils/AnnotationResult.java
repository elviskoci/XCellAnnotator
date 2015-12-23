/**
 * 
 */
package de.tudresden.annotator.annotations.utils;

/**
 * Objects of this class represent the result of the request to create a new range annotation  
 * @author Elvis Koci
 */
public class AnnotationResult {
	
	private ValidationResult validationResult;
	private String message;
				
	/**
	 * @param message
	 * @param statusCode
	 */
	public AnnotationResult(ValidationResult validationResult, String message) {
		this.message = message;
		this.validationResult = validationResult;
	}

	/**
	 * @return the message
	 */
	public String getMessage() {
		return message;
	}

	/**
	 * @return the validationResult
	 */
	public ValidationResult getValidationResult() {
		return validationResult;
	}
		
}
