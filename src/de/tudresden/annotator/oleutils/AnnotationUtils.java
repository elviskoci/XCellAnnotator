/**
 * 
 */
package de.tudresden.annotator.oleutils;

import javax.sql.rowset.CachedRowSet;

import org.eclipse.swt.ole.win32.OleAutomation;

import de.tudresden.annotator.classes.AnnotationClass;
import de.tudresden.annotator.classes.AnnotationTool;

/**
 * @author Elvis Koci
 */
public class AnnotationUtils {
	
	/** 
	 * Annotate the selected areas by drawing a border around each one of them 
	 * 
	 * @param colorIndex the index of the color in the current palette
	 */
	public static void annotateByBorderingSelectedAreas(OleAutomation workbookAutomation, String sheetName, String selectedAreas[], AnnotationClass annotationClass){
		 		
		// get the OleAutomation object for the active worksheet using its name
		OleAutomation worksheetAutomation = WorkbookUtils.getWorksheetAutomationByName(workbookAutomation,sheetName);
	
		// unprotect the worksheet in order to create the textbox
		WorksheetUtils.unprotectWorksheet(worksheetAutomation);
		
		// for each area in the range draw a border
		for (String area : selectedAreas) {
			String[] subStrings = area.split(":");
			
			String topRightCell = subStrings[0];
			String downLeftCell = null; 
			if(subStrings.length==2)
				downLeftCell = subStrings[1];
			
			OleAutomation rangeAutomation = WorksheetUtils.getRangeAutomation(worksheetAutomation, topRightCell, downLeftCell);
			
			long color = annotationClass.getLineColor();
			if(annotationClass.getLineColor()<0){
				color = annotationClass.getColor();
			}
				
			RangeUtils.drawBorderAroundRange(rangeAutomation, annotationClass.getLineStyle(), annotationClass.getLineWeight(), color);	
			rangeAutomation.dispose();
		}
		
		// protect the worksheet to prevent the user from modifying the content of the sheet
		WorksheetUtils.protectWorksheet(worksheetAutomation);
		worksheetAutomation.dispose();
	}
	

	/**
	 * Annotate the selected areas by drawing textbox on top of each one of them.
	 * The color of the textbox depends on the Annotation Class.
	 * 
	 * @param workbookAutomation
	 * @param sheetName
	 * @param selectedAreas
	 * @param annotationClass
	 */
	public static void annotateSelectedAreasWithTextboxes(OleAutomation workbookAutomation, String sheetName, 
																 String selectedAreas[], AnnotationClass annotationClass){
		
		// get the OleAutomation object for the active worksheet using its name
		OleAutomation worksheetAutomation = WorkbookUtils.getWorksheetAutomationByName(workbookAutomation,sheetName);
	
		// unprotect the worksheet in order to create the textbox
		WorksheetUtils.unprotectWorksheet(worksheetAutomation);
		
		// for each area in the range draw a textbox
		for (String area : selectedAreas) {
			String[] subStrings = area.split(":");
			
			String topRightCell = subStrings[0];
			String downLeftCell = null; 
			if(subStrings.length==2)
				downLeftCell = subStrings[1];
			
			OleAutomation rangeAutomation = WorksheetUtils.getRangeAutomation(worksheetAutomation, topRightCell, downLeftCell);
			// get the range positions (location). The textbox will cover the range of cells.
			double left = RangeUtils.getRangeLeftPosition(rangeAutomation);  
			double top = RangeUtils.getRangeTopPosition(rangeAutomation);
			double width = RangeUtils.getRangeWidth(rangeAutomation);
			double height = RangeUtils.getRangeHeight(rangeAutomation);
			
			OleAutomation shapesAutomation = WorksheetUtils.getWorksheetShapes(worksheetAutomation);
			OleAutomation textboxAutomation = ShapeUtils.drawTextBox(shapesAutomation, left, top, width, height); 
			
			setAnnotationProperties(textboxAutomation, annotationClass);
			
			textboxAutomation.dispose();
			shapesAutomation.dispose();
			rangeAutomation.dispose();
		}
		
		// protect the worksheet to prevent user from modifying the annotations
		WorksheetUtils.protectWorksheet(worksheetAutomation);
		worksheetAutomation.dispose();
	}
	
	/**
	 * Annotate the selected areas by drawing textbox on top of each one of them.
	 * The color of the textbox depends on the Annotation Class.
	 * 
	 * @param workbookAutomation
	 * @param sheetName
	 * @param selectedAreas
	 * @param annotationClass
	 */
	public static void annotateSelectedAreasWithRectangle(OleAutomation workbookAutomation, String sheetName, 
																 String selectedAreas[], AnnotationClass annotationClass){
		// get the OleAutomation object for the active worksheet using its name
		OleAutomation worksheetAutomation = WorkbookUtils.getWorksheetAutomationByName(workbookAutomation, sheetName);
	
		// unprotect the worksheet in order to create the textbox
		WorksheetUtils.unprotectWorksheet(worksheetAutomation);
		
		// for each area in the range draw a textbox
		for (String area : selectedAreas) {
			String[] subStrings = area.split(":");
			
			String topRightCell = subStrings[0];
			String downLeftCell = null; 
			if(subStrings.length==2)
				downLeftCell = subStrings[1];
			
			OleAutomation rangeAutomation = WorksheetUtils.getRangeAutomation(worksheetAutomation, topRightCell, downLeftCell);
			// get the range positions (location). The rectangle will surround the range
			double left = RangeUtils.getRangeLeftPosition(rangeAutomation)-2;  
			double top = RangeUtils.getRangeTopPosition(rangeAutomation)-2;
			double width = RangeUtils.getRangeWidth(rangeAutomation)+3.5;
			double height = RangeUtils.getRangeHeight(rangeAutomation)+3.5;
			
			// draw a rectangle around the range
			OleAutomation shapesAutomation = WorksheetUtils.getWorksheetShapes(worksheetAutomation);
			int msoAutoShapeType = 1; // msoShapeRectangle = 1 
			OleAutomation rectangleAutomation = ShapeUtils.drawShape(shapesAutomation, msoAutoShapeType, left, top, width, height);
			
			setAnnotationProperties(rectangleAutomation, annotationClass);
			
			rectangleAutomation.dispose();
			shapesAutomation.dispose();
			rangeAutomation.dispose();
		}
		
		// protect the worksheet to prevent user from modifying the annotations
		WorksheetUtils.protectWorksheet(worksheetAutomation);
		worksheetAutomation.dispose();
	}
	
	
	public static void setAnnotationProperties(OleAutomation annotation, AnnotationClass annotationClass ){
		
		// set fill
		OleAutomation fillFormatAutomation = ShapeUtils.getFillFormatAutomation(annotation);	
		if(annotationClass.hasFill()){
			//TODO: Handle specific cases  
			ColorFormatUtils.setBackColor(fillFormatAutomation, annotationClass.getColor());
			ColorFormatUtils.setForeColor(fillFormatAutomation, annotationClass.getColor());
			FillFormatUtils.setFillTransparency(fillFormatAutomation, annotationClass.getFillTransparency());
		}else{
			FillFormatUtils.setFillVisibility(fillFormatAutomation, false);  	
		}
		
		// set border/line
		OleAutomation lineFormatAutomation = ShapeUtils.getLineFormatAutomation(annotation);
		if(annotationClass.useLine()){					
			LineFormatUtils.setLineStyle(lineFormatAutomation, annotationClass.getLineStyle());
			LineFormatUtils.setLineWeight(lineFormatAutomation, annotationClass.getLineWeight());
			ColorFormatUtils.setForeColor(lineFormatAutomation, annotationClass.getLineColor());
			LineFormatUtils.setLineTransparency(lineFormatAutomation, annotationClass.getLineTransparency());	
			
		}else{
			LineFormatUtils.setLineVisibility(lineFormatAutomation, false);
		}
		
		// set shadow
		OleAutomation shadowFormatAutomation = ShapeUtils.getShadowFormatAutomation(annotation);
		if(annotationClass.useShadow()){
			ShadowFormatUtils.setShadowType(shadowFormatAutomation, annotationClass.getShadowType()); 
			ShadowFormatUtils.setShadowStyle(shadowFormatAutomation, annotationClass.getShadowStyle());
			ShadowFormatUtils.setShadowSize(shadowFormatAutomation, annotationClass.getShadowSize());
			ShadowFormatUtils.setShadowBlur(shadowFormatAutomation, annotationClass.getShadowBlur());
			ColorFormatUtils.setForeColor(shadowFormatAutomation, annotationClass.getColor());
			ShadowFormatUtils.setShadowTransparency(shadowFormatAutomation, annotationClass.getShadowTransparency());				
		}else{
			ShadowFormatUtils.setShadowVisibility(shadowFormatAutomation, false);
		}
	
		if(annotationClass.useText()){		
			OleAutomation textFrameAutomation = ShapeUtils.getTextFrameAutomation(annotation);
			TextFrameUtils.setHorizontalAlignment(textFrameAutomation, annotationClass.getTextHAlignment());  
			TextFrameUtils.setVerticalAlignment(textFrameAutomation, annotationClass.getTextVAlignment());
			
			OleAutomation charactersAutomation = TextFrameUtils.getCharactersAutomation(textFrameAutomation);
			CharactersUtils.setText(charactersAutomation, annotationClass.getText());
			
			OleAutomation fontAutomation = CharactersUtils.getFontAutomation(charactersAutomation);
			FontUtils.setFontColor(fontAutomation, annotationClass.getTextColor());
			FontUtils.setBoldFont(fontAutomation, annotationClass.isBoldText()); 
			FontUtils.setFontSize(fontAutomation, annotationClass.getFontSize()); // TODO: should be relative to the size of the range 
			
			fontAutomation.dispose();
			charactersAutomation.dispose();
			textFrameAutomation.dispose();
		}

		shadowFormatAutomation.dispose();
		lineFormatAutomation.dispose();
		fillFormatAutomation.dispose();
	}
	
	
	
	/**
	 * Call the appropriate annotation function for the given annotation class
	 * @param workbookAutomation
	 * @param sheetName
	 * @param selectedAreas
	 * @param annotationClass
	 */
	public static void callAnnotationMethod(OleAutomation workbookAutomation, String sheetName, 
														String selectedAreas[], AnnotationClass annotationClass ){
		
			 switch (annotationClass.getAnnotationTool()) {
			    case RECTANGLE  : annotateSelectedAreasWithRectangle(workbookAutomation, sheetName, selectedAreas, annotationClass); break;
			    case TEXTBOX  : annotateSelectedAreasWithTextboxes(workbookAutomation, sheetName, selectedAreas, annotationClass); break;
			    case BORDERAROUND: annotateByBorderingSelectedAreas(workbookAutomation, sheetName, selectedAreas, annotationClass); break;
			    default: System.out.println("Option not recognized"); System.exit(1); break;
			}
		
	}

}
