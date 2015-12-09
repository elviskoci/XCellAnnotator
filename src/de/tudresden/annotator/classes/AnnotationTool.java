/**
 * 
 */
package de.tudresden.annotator.classes;

/**
 * @author Elvis Koci
 */
public enum AnnotationTool {
	
	RECTANGLE (1),
	TEXTBOX (2),
	BORDERAROUND (3),
	RANGEFILL (4);
	
	private final int code;

	private AnnotationTool(int itemCode){
		this.code = itemCode;
	}

	/**
	 * @return the code
	 */
	public int getCode() {
		return code;
	}
}
