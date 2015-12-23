/**
 * 
 */
package de.tudresden.annotator.annotations.utils;

/**
 * @author Elvis Koci
 */
public enum ValidationResult {
	
	OK (1), 
	OVERLAPPING (2), 
	NOTCONTAINED (3);

	private final int code;
    
	private ValidationResult(int code) {
		this.code = code;
	}

	/**
	 * @return the status
	 */
	protected int getCode() {
		return code;
	}

}
