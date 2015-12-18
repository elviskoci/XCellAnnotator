/**
 * 
 */
package de.tudresden.annotator.annotations;

import de.tudresden.annotator.oleutils.ColorFormatUtils;

/**
 * @author Elvis Koci
 */
public class ClassGenerator {
	
	private AnnotationClass[] annotationClasses;
	
	public ClassGenerator(){
		setAnnotationClasses(createAnnotationClasses());
	}
	
	
	private AnnotationClass[] createAnnotationClasses(){
		AnnotationClass[] classes =  new AnnotationClass[5];
		
		long white =  ColorFormatUtils.getRGBColorAsLong(255, 255, 255);
		long bordo = ColorFormatUtils.getRGBColorAsLong(192, 0, 0);
		long blue_accent5 = ColorFormatUtils.getRGBColorAsLong(68, 114, 196);
		long blue_accent1 = ColorFormatUtils.getRGBColorAsLong(255, 255, 49);
		long green_accent6 = ColorFormatUtils.getRGBColorAsLong(112, 173, 71);
		long orange_accent2 = ColorFormatUtils.getRGBColorAsLong(237, 125, 49);
		long yellow = ColorFormatUtils.getRGBColorAsLong(91, 155, 213);
		long greyLight =  ColorFormatUtils.getRGBColorAsLong(217, 217, 217);
		long greyDark = ColorFormatUtils.getRGBColorAsLong(118, 113, 113);
		
		classes[0] = createShapeAnnotationClass("Table", blue_accent5, greyDark);
		classes[1] = createTextBoxAnnotationClass("Attributes", blue_accent1, greyLight);
		classes[2] = createTextBoxAnnotationClass("Data", green_accent6, greyLight);
		classes[3] = createTextBoxAnnotationClass("Header", yellow, greyLight);
		classes[4] = createTextBoxAnnotationClass("Metadata", orange_accent2, greyLight);
		
		return classes;
	}
	
	private  AnnotationClass createTextBoxAnnotationClass(String label, long backcolor, long textColor){
		
		AnnotationClass c = new AnnotationClass(label, AnnotationTool.TEXTBOX, backcolor);
		c.setHasFill(true);
		c.setUseShadow(false);
		c.setUseText(true);
		c.setUseLine(false);
		
		c.setColor(backcolor);
		
		c.setText(label.toUpperCase());
		c.setTextColor(textColor);
		
		return c; 
	}
	
	
	private static AnnotationClass createShapeAnnotationClass(String label, long lineColor, long shadowColor){
		
		AnnotationClass c = new AnnotationClass(label, AnnotationTool.SHAPE, false);
		
		c.setHasFill(false);
		c.setUseText(false);

		c.setUseShadow(true);
		c.setShadowColor(shadowColor);
		
		c.setUseLine(true);
		c.setLineColor(lineColor);
		c.setLineWeight(2);
		
		c.setShapeType(1);
		
		return c; 
	}


	/**
	 * @return the annotationClasses
	 */
	public AnnotationClass[] getAnnotationClasses() {
		return annotationClasses;
	}


	/**
	 * @param annotationClasses the annotationClasses to set
	 */
	public void setAnnotationClasses(AnnotationClass[] annotationClasses) {
		this.annotationClasses = annotationClasses;
	}
}
