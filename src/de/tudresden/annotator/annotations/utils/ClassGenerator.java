/**
 * 
 */
package de.tudresden.annotator.annotations.utils;

import java.util.HashMap;

import de.tudresden.annotator.annotations.AnnotationClass;
import de.tudresden.annotator.annotations.AnnotationTool;
import de.tudresden.annotator.oleutils.ColorFormatUtils;

/**
 * @author Elvis Koci
 */
public class ClassGenerator {
	
	private static final HashMap<String, AnnotationClass> annotationClasses;
	
	static{
		annotationClasses = new HashMap<String, AnnotationClass>();
		AnnotationClass[] classes = createAnnotationClasses();
		for (AnnotationClass annotationClass : classes) {
			annotationClasses.put(annotationClass.getLabel(), annotationClass);
		}
	}
		
	private static AnnotationClass[] createAnnotationClasses(){
		
		AnnotationClass[] classes =  new AnnotationClass[5];
		
		// long white =  ColorFormatUtils.getRGBColorAsLong(255, 255, 255);
		// long bordo = ColorFormatUtils.getRGBColorAsLong(192, 0, 0);
		long blue_accent5 = ColorFormatUtils.getRGBColorAsLong(68, 114, 196);
		long blue_accent1 = ColorFormatUtils.getRGBColorAsLong(255, 255, 49);
		long green_accent6 = ColorFormatUtils.getRGBColorAsLong(112, 173, 71);
		long orange_accent2 = ColorFormatUtils.getRGBColorAsLong(237, 125, 49);
		long yellow = ColorFormatUtils.getRGBColorAsLong(91, 155, 213);
		long greyLight =  ColorFormatUtils.getRGBColorAsLong(217, 217, 217);
		long greyDark = ColorFormatUtils.getRGBColorAsLong(118, 113, 113);
		
		// table can contains all the other classes
		classes[0] = createShapeAnnotationClass("Table", 1, 0, true, blue_accent5, 2, true, greyDark, true, false, false, null); 
		classes[1] = createTextBoxAnnotationClass("Attributes", blue_accent1, true, greyLight, true, true, classes[0]);
		classes[2] = createTextBoxAnnotationClass("Data", green_accent6,  true, greyLight, true, true, classes[0]);
		classes[3] = createTextBoxAnnotationClass("Header", yellow,  true, greyLight, true, true, classes[0]);
		// metadata can be outside of a table or inside. Tables can share metadata
		classes[4] = createTextBoxAnnotationClass("Metadata", orange_accent2, true, greyLight, true, false, null); 
		
		return classes;
	}
	
	
	private static AnnotationClass createShapeAnnotationClass(String label, int shapeType, long fillColor, 
													boolean useLine, long lineColor, int lineWeight, boolean useShadow, long shadowColor, 
												    boolean isContainer, boolean isContainable, boolean isDependent, AnnotationClass container){
		
		AnnotationClass c = new AnnotationClass(label, AnnotationTool.SHAPE, false);
		
		if(isContainer){
			c.setHasFill(false);
		}else{
			c.setHasFill(true);
		}
		c.setColor(fillColor);
		
		
		c.setUseShadow(useShadow);
		c.setShadowColor(shadowColor);
		
		c.setUseLine(useLine);
		c.setLineColor(lineColor);
		c.setLineWeight(lineWeight);
		
		c.setShapeType(shapeType);
		
		c.setUseText(false);
		
		c.setIsContainer(isContainer);
		c.setCanBeContained(isContainable);
		c.setIsDependent(isDependent);
		c.setContainer(container);
			
		return c; 
	}

	private  static AnnotationClass createTextBoxAnnotationClass(String label, long backcolor, boolean useText, long textColor,
														    boolean isContainable, boolean isDependent, AnnotationClass container){
		
		AnnotationClass c = new AnnotationClass(label, AnnotationTool.TEXTBOX, backcolor);
		
		c.setHasFill(true);
		c.setColor(backcolor);
		
		c.setUseShadow(false);
		c.setUseLine(false);
		
		c.setUseText(useText);
		c.setText(label.toUpperCase());
		c.setTextColor(textColor);
		
		c.setCanBeContained(isContainable);
		c.setIsDependent(isDependent);
		c.setContainer(container);
		
		return c; 
	}


	/**
	 * @return the annotationclasses
	 */
	public static HashMap<String, AnnotationClass> getAnnotationClasses() {
		return annotationClasses;
	}
	
}
