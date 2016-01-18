/**
 * 
 */
package de.tudresden.annotator.annotations.utils;

/**
 * @author Elvis Koci
 */
public enum ValidationResult {
	
	OK (1),
	EMPTY (2),
	OVERLAPPING (3), 
	NOTCONTAINED (4),
	NOSELECTION (5);

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
