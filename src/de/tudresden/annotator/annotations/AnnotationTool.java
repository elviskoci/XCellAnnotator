/**
 * 
 */
package de.tudresden.annotator.annotations;

/**
 * @author Elvis Koci
 */
public enum AnnotationTool {
	
	BORDERAROUND (1),
	RANGEFILL (2),
	TEXTBOX (3),
	SHAPE (4),
	COMPLEXSHAPE (5);
	
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
