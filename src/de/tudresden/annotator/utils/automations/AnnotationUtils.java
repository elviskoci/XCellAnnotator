/**
 * 
 */
package de.tudresden.annotator.utils.automations;

import org.eclipse.swt.ole.win32.OleAutomation;

/**
 * @author Elvis Koci
 */
public class AnnotationUtils {
	
	/** 
	 * Annotate the selected areas by drawing a border around each one of them 
	 * 
	 * @param colorIndex the index of the color in the current palette
	 */
	public static void annotateByBorderingSelectedAreas(OleAutomation workbookAutomation, String sheetName, String selectedAreas[], int colorIndex){
		 		
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
			
			RangeUtils.drawBorderAroundRange(rangeAutomation,1,4,colorIndex);
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
	 * @param color a long that represents a RGB color. Is calculated as B * 65536 + G * 256 + R
	 * @param label
	 */
	public static void annotateSelectedAreasWithTextboxes(OleAutomation workbookAutomation, String sheetName, String selectedAreas[], long color, String label){
		
		// get the OleAutomation object for the active worksheet using its name
		OleAutomation worksheetAutomation = WorkbookUtils.getWorksheetAutomationByName(workbookAutomation,sheetName);
	
		// unprotect the worksheet in order to create the textbox
		// WorksheetUtils.unprotectWorksheet(worksheetAutomation);
		
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
			
			OleAutomation fillFormatAutomation = ShapeUtils.getFillFormatAutomation(textboxAutomation);
			FillFormatUtils.setFillTransparency(fillFormatAutomation, 0.80);
			ColorFormatUtils.setForeColor(fillFormatAutomation, color);
			
			OleAutomation lineFormatAutomation = ShapeUtils.getLineFormatAutomation(textboxAutomation);
			LineFormatUtils.setLineVisibility(lineFormatAutomation, false);
			
			OleAutomation textFrameAutomation = ShapeUtils.getTextFrameAutomation(textboxAutomation);
			TextFrameUtils.setHorizontalAlignment(textFrameAutomation, -4108); // align center 
			TextFrameUtils.setVerticalAlignment(textFrameAutomation, -4108); // align center
			
			OleAutomation charactersAutomation = TextFrameUtils.getCharactersAutomation(textFrameAutomation);
			CharactersUtils.setText(charactersAutomation, label);
			
			OleAutomation fontAutomation = CharactersUtils.getFontAutomation(charactersAutomation);
			long whiteColor = 255 * 65536 + 255 * 256 + 255;
			FontUtils.setFontColor(fontAutomation, whiteColor);
			FontUtils.setBoldFont(fontAutomation, true); 
			//FontUtils.setFontSize(fontAutomation, 11); // TODO: should be relative to the size of the range 
			
			fontAutomation.dispose();
			charactersAutomation.dispose();
			textFrameAutomation.dispose();
			
			fillFormatAutomation.dispose();
			textboxAutomation.dispose();
			shapesAutomation.dispose();
			rangeAutomation.dispose();
		}
		
		// protect the worksheet to prevent user from modifying the annotations
		// WorksheetUtils.protectWorksheet(worksheetAutomation);
		worksheetAutomation.dispose();
	}
	
	/**
	 * Annotate the selected areas by drawing textbox on top of each one of them.
	 * The color of the textbox depends on the Annotation Class.
	 * 
	 * @param workbookAutomation
	 * @param sheetName
	 * @param selectedAreas
	 * @param color a long that represents a RGB color. Is calculated as B * 65536 + G * 256 + R
	 * @param label
	 */
	public static void annotateSelectedAreasWithRectangle(OleAutomation workbookAutomation, String sheetName, String selectedAreas[]){
		
		// get the OleAutomation object for the active worksheet using its name
		OleAutomation worksheetAutomation = WorkbookUtils.getWorksheetAutomationByName(workbookAutomation, sheetName);
	
		// unprotect the worksheet in order to create the textbox
		// WorksheetUtils.unprotectWorksheet(worksheetAutomation);
		
		// for each area in the range draw a textbox
		for (String area : selectedAreas) {
			String[] subStrings = area.split(":");
			
			String topRightCell = subStrings[0];
			String downLeftCell = null; 
			if(subStrings.length==2)
				downLeftCell = subStrings[1];
			
			OleAutomation rangeAutomation = WorksheetUtils.getRangeAutomation(worksheetAutomation, topRightCell, downLeftCell);
			// get the range positions (location). The rectangle will surround the range
			double left = RangeUtils.getRangeLeftPosition(rangeAutomation)-1;  
			double top = RangeUtils.getRangeTopPosition(rangeAutomation)-1;
			double width = RangeUtils.getRangeWidth(rangeAutomation)+2;
			double height = RangeUtils.getRangeHeight(rangeAutomation)+2;
			
			// draw a rectangle around the range
			OleAutomation shapesAutomation = WorksheetUtils.getWorksheetShapes(worksheetAutomation);
			int msoAutoShapeType = 1; // msoShapeRectangle = 1 
			OleAutomation rectangleAutomation = ShapeUtils.drawShape(shapesAutomation, msoAutoShapeType, left, top, width, height);
			
			// set no fill for the rectangle
			OleAutomation fillFormatAutomation = ShapeUtils.getFillFormatAutomation(rectangleAutomation);
			FillFormatUtils.setVisible(fillFormatAutomation, false);  	
			
			
			long black = ColorFormatUtils.getRGBColorAsLong(0, 0, 0);
			long blue = ColorFormatUtils.getRGBColorAsLong(65, 113, 156);
			long bordo = ColorFormatUtils.getRGBColorAsLong(192, 0, 0);
			long grey = ColorFormatUtils.getRGBColorAsLong(118, 113, 113);
			
			// set the color and weight of the border for the rectangle shape
			OleAutomation lineFormatAutomation = ShapeUtils.getLineFormatAutomation(rectangleAutomation);
			LineFormatUtils.setLineVisibility(lineFormatAutomation, true);
			LineFormatUtils.setLineWeight(lineFormatAutomation, 1.5 );
			ColorFormatUtils.setForeColor(lineFormatAutomation, bordo);
			
			// set shadow around the rectangle
			OleAutomation shadowFormatAutomation = ShapeUtils.getShadowFormatAutomation(rectangleAutomation);
			ShadowFormatUtils.setShadowVisibility(shadowFormatAutomation, true);
			ShadowFormatUtils.setShadowBlur(shadowFormatAutomation, 4);
			ShadowFormatUtils.setShadowSize(shadowFormatAutomation, 100);
			ShadowFormatUtils.setShadowStyle(shadowFormatAutomation, 2); // outer shadow
			ShadowFormatUtils.setShadowTransparency(shadowFormatAutomation, 0.30);
			ShadowFormatUtils.setShadowOffsetX(shadowFormatAutomation, 0);
			ShadowFormatUtils.setShadowOffsetY(shadowFormatAutomation, 0);	
			ColorFormatUtils.setForeColor(shadowFormatAutomation, grey);
			
			shadowFormatAutomation.dispose();
			fillFormatAutomation.dispose();
			rectangleAutomation.dispose();
			shapesAutomation.dispose();
			rangeAutomation.dispose();
		}
		
		// protect the worksheet to prevent user from modifying the annotations
		//WorksheetUtils.protectWorksheet(worksheetAutomation);
		worksheetAutomation.dispose();
	}

}
