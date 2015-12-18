/**
 * 
 */
package de.tudresden.annotator.annotations;

import java.util.HashSet;

import org.eclipse.swt.ole.win32.OleAutomation;

/**
 * @author Elvis Koci
 */
public class AnnotationClass {
	
	/*
	 * Experimental 
	 * An automation that represents a template to be used for annotation. 
	 * It can be a shape which attributes are set (text,line,shadow,fill, etc) 
	 */
	private OleAutomation annotationAutomation;
	private int priority = -1;
	private boolean isContainer = false;
	private boolean isDependent = false;
	private HashSet<AnnotationClass>  containedClasses = new HashSet<AnnotationClass>();
	private String name;

	/*
	 * The label (name) used for the annotation class 
	 */
	private String label; 
	
	/*
	 * The annotation tool that will be used to annotate 
	 */
	private AnnotationTool annotationTool; 
	
	
	/*
	 * The color associated with this annotation class.
	 * It is an RGB color represented as long using this formula: B * 65536 + G * 256 + R
	 */
	private long color;
	
	/*
	 * Set if to use fill or not. If annotation tool is a shape, this property will apply to shape fill. 
	 * If a simple border is used to annotate, than this property it is not considered.   
	 */
	private boolean hasFill = true; 
	private double fillTransparency = 0.80;
	
	private boolean useShadow = false;
	private int shadowType = 21; // offset diagonal bottom right
	private int shadowStyle = 2; // outer shadow
	private int shadowBlur = 5;
	private long shadowColor = -1;
	private int shadowSize = 100;
	private double shadowTransparency = 0.45;
	
	private boolean useText = false;
	private String text = null;
	private long textColor = -1;
	private boolean boldText = true;
	private int fontSize = 11;
	private int textHAlignment = -4108; // align center
	private int textVAlignment = -4108; // align center
	
	private boolean useLine = false;
	private double lineWeight = 1;
	private long lineColor = -1; 
	private int lineStyle = 1; // Single line
	private double lineTransparency = 0;
	
	private int shapeType = 1; //default rectangle
	
	/**
	 * 
	 * @param classLabel a string used as label (name) for the annotation class
	 * @param tool the annotation tool that will be used to annotate this class 
	 * @param color an RGB color represented as long using this formula: B * 65536 + G * 256 + R
	 */
	public AnnotationClass(String classLabel, AnnotationTool tool, long color) {
		
		this.label = classLabel;
		this.color = color;
		this.annotationTool = tool;
	}
	
	
	/**
	 * 
	 * @param classLabel a string used as label (name) for the annotation class
	 * @param tool the annotation tool that will be used to annotate this class 
	 * @param hasFill set if the annotation tool has or not a fill
	 */
	public AnnotationClass(String classLabel, AnnotationTool tool, boolean hasFill) {
		
		this.label = classLabel;
		this.annotationTool = tool;
		this.hasFill = hasFill;
	}
	
	
	/**
	 * 
	 * @param classLabel a string used as label (name) for the annotation class
	 * @param tool the annotation tool that will be used to annotate this class 
	 * @param hasFill set if the annotation tool has or not a fill
	 * @param color an RGB color represented as long using this formula: B * 65536 + G * 256 + R
	 */
	public AnnotationClass(String classLabel, AnnotationTool tool, boolean hasFill , long color) {
		
		this.label = classLabel;
		this.color = color;
		this.annotationTool = tool;
		this.hasFill = hasFill;
	}
	
	/**
	 * 
	 * @param classLabel a string used as label (name) for the annotation class
	 * @param tool the annotation tool that will be used to annotate this class 
	 * @param shapeType the type of AutoShape to create
	 * @param hasFill set if the annotation tool has or not a fill
	 */
	public AnnotationClass(String classLabel, AnnotationTool tool, int shapeType, boolean hasFill) {
		
		this.label = classLabel;
		this.shapeType = shapeType;
		this.annotationTool = tool;
		this.hasFill = hasFill;
	}
	
	
	/**
	 * 
	 * @param classLabel a string used as label (name) for the annotation class
	 * @param tool the annotation tool that will be used to annotate this class 
	 * @param shapeType the type of AutoShape to create
	 * @param hasFill set if the annotation tool has or not a fill
	 * @param color an RGB color represented as long using this formula: B * 65536 + G * 256 + R
	 */
	public AnnotationClass(String classLabel, AnnotationTool tool, int shapeType, boolean hasFill , long color) {
		
		this.label = classLabel;
		this.color = color;
		this.annotationTool = tool;
		this.hasFill = hasFill;
		this.shapeType = shapeType;
	}
	
	
	/**
	 * 
	 * @param lineStyle
	 * @param lineWeight
	 * @param lineColor
	 * @param lineTransparency
	 */
	public void setLineProperties( int lineStyle, double lineWeight, long lineColor, double lineTransparency){
		
		this.lineStyle = lineStyle;
		this.lineColor = lineColor;
		this.lineWeight = lineWeight;
		this.lineTransparency = lineTransparency;
	}
	
	/**
	 * 
	 * @param shadowType
	 * @param shadowStyle
	 * @param shadowBlur
	 * @param shadowColor
	 * @param shadowSize
	 * @param shadowTransparency
	 */
	public void setShadowProperties(int shadowType, int shadowStyle, int shadowBlur, long shadowColor, int shadowSize, double shadowTransparency){
		
		this.shadowType = shadowType;
		this.shadowBlur = shadowBlur;
		this.shadowSize = shadowSize;
		this.shadowStyle = shadowStyle;
		this.shadowColor = shadowColor;
		this.shadowTransparency = shadowTransparency;
	}
	
	/**
	 * 
	 * @param text
	 * @param textColor
	 * @param boldText
	 * @param fontSize
	 * @param textHAlignment
	 * @param textVAlignment
	 */
    public void setTextProperties(String text, long textColor, boolean boldText, int fontSize, int textHAlignment, int textVAlignment ){
    	
    	this.text = text;
    	this.textColor = textColor;
    	this.boldText = boldText;
    	this.fontSize = fontSize;
    	this.textHAlignment = textHAlignment;
    	this.textVAlignment = textVAlignment;
    }
	
   
	/**
	 * @return the label
	 */
	public String getLabel() {
		return label;
	}

	/**
	 * @param label the label to set
	 */
	public void setLabel(String label) {
		this.label = label;
	}

	/**
	 * @return the color
	 */
	public long getColor() {
		return color;
	}

	/**
	 * @param color the color to set
	 */
	public void setColor(long color) {
		this.color = color;
	}

	/**
	 * @return the annotationTool
	 */
	public AnnotationTool getAnnotationTool() {
		return annotationTool;
	}

	/**
	 * @param annotationTool the annotationTool to set
	 */
	public void setAnnotationTool(AnnotationTool annotationTool) {
		this.annotationTool = annotationTool;
	}
	
	/**
	 * @return the shapeType
	 */
	public int getShapeType() {
		return shapeType;
	}

	/**
	 * @param shapeType the shapeType to set
	 */
	public void setShapeType(int shapeType) {
		this.shapeType = shapeType;
	}


	/**
	 * @return the useShadow
	 */
	public boolean useShadow() {
		return useShadow;
	}

	/**
	 * @param useShadow the useShadow to set
	 */
	public void setUseShadow(boolean useShadow) {
		this.useShadow = useShadow;
	}

	/**
	 * @return the shadowType
	 */
	public int getShadowType() {
		return shadowType;
	}

	/**
	 * @param shadowType the shadowType to set
	 */
	public void setShadowType(int shadowType) {
		this.shadowType = shadowType;
	}

	/**
	 * @return the shadowStyle
	 */
	public int getShadowStyle() {
		return shadowStyle;
	}

	/**
	 * @param shadowStyle the shadowStyle to set
	 */
	public void setShadowStyle(int shadowStyle) {
		this.shadowStyle = shadowStyle;
	}

	/**
	 * @return the shadowBlur
	 */
	public int getShadowBlur() {
		return shadowBlur;
	}

	/**
	 * @param shadowBlur the shadowBlur to set
	 */
	public void setShadowBlur(int shadowBlur) {
		this.shadowBlur = shadowBlur;
	}

	/**
	 * @return the shadowColor
	 */
	public long getShadowColor() {
		return shadowColor;
	}

	/**
	 * @param shadowColor the shadowColor to set
	 */
	public void setShadowColor(long shadowColor) {
		this.shadowColor = shadowColor;
	}

	/**
	 * @return the shadowSize
	 */
	public int getShadowSize() {
		return shadowSize;
	}

	/**
	 * @param shadowSize the shadowSize to set
	 */
	public void setShadowSize(int shadowSize) {
		this.shadowSize = shadowSize;
	}

	/**
	 * @return the shadowTransparency
	 */
	public double getShadowTransparency() {
		return shadowTransparency;
	}

	/**
	 * @param shadowTransparency the shadowTransparency to set
	 */
	public void setShadowTransparency(double shadowTransparency) {
		this.shadowTransparency = shadowTransparency;
	}

	/**
	 * @return the useText
	 */
	public boolean useText() {
		return useText;
	}

	/**
	 * @param useText the useText to set
	 */
	public void setUseText(boolean useText) {
		this.useText = useText;
	}

	/**
	 * @return the text
	 */
	public String getText() {
		return text;
	}

	/**
	 * @param text the text to set
	 */
	public void setText(String text) {
		this.text = text;
	}

	/**
	 * @return the textColor
	 */
	public long getTextColor() {
		return textColor;
	}

	/**
	 * @param textColor the textColor to set
	 */
	public void setTextColor(long textColor) {
		this.textColor = textColor;
	}

	/**
	 * @return the boldText
	 */
	public boolean isBoldText() {
		return boldText;
	}

	/**
	 * @param boldText the boldText to set
	 */
	public void setBoldText(boolean boldText) {
		this.boldText = boldText;
	}

	/**
	 * @return the fontSize
	 */
	public int getFontSize() {
		return fontSize;
	}

	/**
	 * @param fontSize the fontSize to set
	 */
	public void setFontSize(int fontSize) {
		this.fontSize = fontSize;
	}

	/**
	 * @return the useLine
	 */
	public boolean useLine() {
		return useLine;
	}

	/**
	 * @param useLine the useLine to set
	 */
	public void setUseLine(boolean useLine) {
		this.useLine = useLine;
	}

	/**
	 * @return the lineWeight
	 */
	public double getLineWeight() {
		return lineWeight;
	}

	/**
	 * @param lineWeight the lineWeight to set
	 */
	public void setLineWeight(double lineWeight) {
		this.lineWeight = lineWeight;
	}

	/**
	 * @return the lineColor
	 */
	public long getLineColor() {
		return lineColor;
	}

	/**
	 * @param lineColor the lineColor to set
	 */
	public void setLineColor(long lineColor) {
		this.lineColor = lineColor;
	}

	/**
	 * @return the lineStyle
	 */
	public int getLineStyle() {
		return lineStyle;
	}

	/**
	 * @param lineStyle the lineStyle to set
	 */
	public void setLineStyle(int lineStyle) {
		this.lineStyle = lineStyle;
	}

	/**
	 * @return the lineTransparency
	 */
	public double getLineTransparency() {
		return lineTransparency;
	}

	/**
	 * @param lineTransparency the lineTransparency to set
	 */
	public void setLineTransparency(double lineTransparency) {
		this.lineTransparency = lineTransparency;
	}


	/**
	 * @return the textHAlignment
	 */
	public int getTextHAlignment() {
		return textHAlignment;
	}


	/**
	 * @param textHAlignment the textHAlignment to set
	 */
	public void setTextHAlignment(int textHAlignment) {
		this.textHAlignment = textHAlignment;
	}


	/**
	 * @return the textVAlignment
	 */
	public int getTextVAlignment() {
		return textVAlignment;
	}


	/**
	 * @param textVAlignment the textVAlignment to set
	 */
	public void setTextVAlignment(int textVAlignment) {
		this.textVAlignment = textVAlignment;
	}


	/**
	 * @return the hasFill
	 */
	public boolean hasFill() {
		return hasFill;
	}


	/**
	 * @param hasFill the hasFill to set
	 */
	public void setHasFill(boolean hasFill) {
		this.hasFill = hasFill;
	}


	/**
	 * @return the fillTransparency
	 */
	public double getFillTransparency() {
		return fillTransparency;
	}


	/**
	 * @param fillTransparency the fillTransparency to set
	 */
	public void setFillTransparency(double fillTransparency) {
		this.fillTransparency = fillTransparency;
	}
}
