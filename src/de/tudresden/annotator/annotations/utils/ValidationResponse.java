/**
 * 
 */
package de.tudresden.annotator.annotations.utils;

/**
 * Objects of this class represent the result from the validation of the request to create a new range annotation
 * @author Elvis Koci
 */
public class ValidationResponse {
	
	private ValidationResult validationResult; 
	private String message;
				
	/**
	 * @param message
	 * @param statusCode
	 */
	public ValidationResponse(ValidationResult statusCode, String message) {
		this.message = message;
		this.validationResult = statusCode;
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
