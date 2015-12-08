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
			// adjust the positions such that textbox is inside the range, but does not cover (hide) its borders
			double left = RangeUtils.getRangeLeftPosition(rangeAutomation) + 0.5;  
			double top = RangeUtils.getRangeTopPosition(rangeAutomation) + 0.5;
			double width = RangeUtils.getRangeWidth(rangeAutomation) - 1.0;
			double height = RangeUtils.getRangeHeight(rangeAutomation) - 1.0;
			
			OleAutomation shapesAutomation = WorksheetUtils.getWorksheetShapes(worksheetAutomation);
			OleAutomation textboxAutomation = ShapeUtils.drawTextBox(shapesAutomation, left, top, width, height); 
			
			OleAutomation fillFormatAutomation = ShapeUtils.getFillFormatAutomation(textboxAutomation);
			FillFormatUtils.setBackgroundColor(fillFormatAutomation, color);
			FillFormatUtils.setFillTransparency(fillFormatAutomation, 0.60);
			
			OleAutomation textFrameAutomation = ShapeUtils.getTextFrameAutomation(textboxAutomation);
			TextFrameUtils.setHorizontalAlignment(textFrameAutomation, -4108); // align center 
			TextFrameUtils.setVerticalAlignment(textFrameAutomation, -4108); // align center
			
			OleAutomation charactersAutomation = TextFrameUtils.getCharactersAutomation(textFrameAutomation);
			CharactersUtils.setText(charactersAutomation, label);
			
			OleAutomation fontAutomation = CharactersUtils.getFontAutomation(charactersAutomation);
			long whiteColor = 255 * 65536 + 255 * 256 + 255;
			FontUtils.setFontColor(fontAutomation, whiteColor);
			FontUtils.setBoldFont(fontAutomation, true); 
			FontUtils.setFontSize(fontAutomation, 11); // TODO: should be relative to the size of the range 
			
			fontAutomation.dispose();
			charactersAutomation.dispose();
			textFrameAutomation.dispose();
			fillFormatAutomation.dispose();
			textboxAutomation.dispose();
			shapesAutomation.dispose();
			rangeAutomation.dispose();
		}
		
		// protect the worksheet to prevent user from modifying the annotations
		WorksheetUtils.protectWorksheet(worksheetAutomation);
		worksheetAutomation.dispose();
	}

}
