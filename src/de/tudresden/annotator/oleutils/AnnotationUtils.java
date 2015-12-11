/**
 * 
 */
package de.tudresden.annotator.oleutils;

import java.util.HashMap;

import org.eclipse.swt.ole.win32.OleAutomation;

import de.tudresden.annotator.annotations.AnnotationClass;

/**
 * @author Elvis Koci
 */
public class AnnotationUtils {
	
	private static final String AnnotationDataSheetName = "Annotation_Data_Sheet";
	private static HashMap<String , Integer> annotationMap = new HashMap<String , Integer>();
	
	/** 
	 * @deprecated
	 * 
	 * Annotate the selected areas by drawing a border around each one of them 
	 * 
	 * @param workbookAutomation
	 * @param sheetName
	 * @param sheetIndex
	 * @param selectedAreas
	 * @param annotationClass
	 */
	public static void annotateByBorderingSelectedAreas(OleAutomation workbookAutomation, String sheetName, int sheetIndex, 
																		String selectedAreas[], AnnotationClass annotationClass){
		 		
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
	 * @param sheetIndex
	 * @param selectedAreas
	 * @param annotationClass
	 */
	public static void annotateSelectedAreasWithTextboxes(OleAutomation workbookAutomation, String sheetName, int sheetIndex,  
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
			
			// get the range positions (location). The textbox will cover the range of cells.
			OleAutomation rangeAutomation = WorksheetUtils.getRangeAutomation(worksheetAutomation, topRightCell, downLeftCell);
			double left = RangeUtils.getRangeLeftPosition(rangeAutomation);  
			double top = RangeUtils.getRangeTopPosition(rangeAutomation);
			double width = RangeUtils.getRangeWidth(rangeAutomation);
			double height = RangeUtils.getRangeHeight(rangeAutomation);
			rangeAutomation.dispose();
			
			// draw the textbox 
			OleAutomation shapesAutomation = WorksheetUtils.getWorksheetShapes(worksheetAutomation);
			OleAutomation textboxAutomation = ShapeUtils.drawTextBox(shapesAutomation, left, top, width, height); 
			setAnnotationProperties(textboxAutomation, annotationClass);
			
			int annotationIndex = getAnnotationIndex(annotationClass.getLabel());
			String annotationObjectName = sheetName.replace(" ", "_")+"_Annotation_"+annotationClass.getLabel()+"_"+annotationIndex;  
			ShapeUtils.setShapeName(textboxAutomation, annotationObjectName);
			
			shapesAutomation.dispose();
			textboxAutomation.dispose();
			
			// save metadata about the annotation
			saveAnnotationData(workbookAutomation, sheetName, sheetIndex, annotationClass.getLabel(), annotationIndex, area);
		}
		
		// protect the worksheet to prevent user from modifying the annotations
		// WorksheetUtils.protectWorksheet(worksheetAutomation);
		worksheetAutomation.dispose();
	}
	
	/**
	 * Annotate the selected areas by drawing a rectangle on top of each one of them.
	 * The color of the rectangle depends on the Annotation Class.
	 * 
	 * @param workbookAutomation
	 * @param sheetName
	 * @param sheetIndex
	 * @param selectedAreas
	 * @param annotationClass
	 */
	public static void annotateSelectedAreasWithRectangle(OleAutomation workbookAutomation, String sheetName, int sheetIndex,
																 String selectedAreas[], AnnotationClass annotationClass){
		
		// get the OleAutomation object for the active worksheet using its name
		OleAutomation worksheetAutomation = WorkbookUtils.getWorksheetAutomationByIndex(workbookAutomation, sheetIndex);
		
		// unprotect the worksheet in order to create the annotation
		WorksheetUtils.unprotectWorksheet(worksheetAutomation);
		
		// annotate each area
		for (String area : selectedAreas) {
			
			String[] subStrings = area.split(":");
			String topRightCell = subStrings[0];
			String downLeftCell = null; 
			if(subStrings.length==2)
				downLeftCell = subStrings[1];
			
			// get the range (area) position. 
			OleAutomation rangeAutomation = WorksheetUtils.getRangeAutomation(worksheetAutomation, topRightCell, downLeftCell);
			double left = RangeUtils.getRangeLeftPosition(rangeAutomation)-1;  
			double top = RangeUtils.getRangeTopPosition(rangeAutomation)-1;
			double width = RangeUtils.getRangeWidth(rangeAutomation)+2;
			double height = RangeUtils.getRangeHeight(rangeAutomation)+2;
			rangeAutomation.dispose();
			
			// draw the shape 
			OleAutomation shapesAutomation = WorksheetUtils.getWorksheetShapes(worksheetAutomation);	
			int msoAutoShapeType = 1; // msoShapeRectangle = 1 
			OleAutomation rectangleAutomation = ShapeUtils.drawShape(shapesAutomation, msoAutoShapeType, left, top, width, height);
			setAnnotationProperties(rectangleAutomation, annotationClass);
			
			int annotationIndex = getAnnotationIndex(annotationClass.getLabel());
			String annotationObjectName = sheetName.replace(" ", "_")+"_Annotation_"+annotationClass.getLabel()+"_"+annotationIndex;  
			ShapeUtils.setShapeName(rectangleAutomation, annotationObjectName);
			
			shapesAutomation.dispose();
			rectangleAutomation.dispose();
				
			// save metadata about the annotation
			saveAnnotationData(workbookAutomation, sheetName, sheetIndex, annotationClass.getLabel(), annotationIndex, area);	
		}
		
		// protect the worksheet to prevent user from modifying the annotations
		// WorksheetUtils.makeWorksheetActive(worksheetAutomation);
		//System.out.println(WorksheetUtils.protectWorksheet(worksheetAutomation));
		worksheetAutomation.dispose();
	}
	
	
	/**
	 * Format the annotation object (shape, textbox, etc) used to annotate 
	 * @param annotation
	 * @param annotationClass
	 */
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
	 * 
	 * @param workbookAutomation
	 * @param sheetName
	 * @param sheetIndex
	 * @param annotationLabel
	 * @param annotationIndex
	 * @param range
	 */
	public static void saveAnnotationData(OleAutomation workbookAutomation, String sheetName, int sheetIndex, String annotationLabel,
																						int annotationIndex , String range){
		
		OleAutomation annotationDataSheet =  WorkbookUtils.getWorksheetAutomationByName(workbookAutomation, AnnotationDataSheetName);
		
		if(annotationDataSheet==null){		
			annotationDataSheet = createAnnotationDataSheet(workbookAutomation);			
		}
		
		OleAutomation usedRange = WorksheetUtils.getUsedRange(annotationDataSheet);		
		String usedAddress = RangeUtils.getRangeAddress(usedRange);
		
		usedAddress = usedAddress.replace("$", "");
		String[] cells =  usedAddress.split(":");
		
		char startColumn =   cells[0].charAt(0); 
		int endRow =   Integer.valueOf(cells[1].substring(1));
		endRow =  endRow + 1;
					
		writeNewRow( annotationDataSheet, startColumn, endRow, sheetName, sheetIndex, annotationLabel, annotationIndex, range);
				
	}
	
	/**
	 * 
	 * @param annotationDataSheet
	 * @param startColumn
	 * @param endRow
	 * @param sheetName
	 * @param sheetIndex
	 * @param label
	 * @param index
	 * @param range
	 */
	public static void writeNewRow(OleAutomation annotationDataSheet, char startColumn, int endRow, String sheetName, 
																int sheetIndex, String label, int index, String range){		
		
		WorksheetUtils.unprotectWorksheet(annotationDataSheet);
		
		OleAutomation cell1 = WorksheetUtils.getRangeAutomation(annotationDataSheet, startColumn+""+endRow, null);
		RangeUtils.setValue(cell1, sheetName);	 
		OleAutomation cell2 = WorksheetUtils.getRangeAutomation(annotationDataSheet, ((char) (startColumn+1))+""+endRow, null);
		RangeUtils.setValue(cell2, String.valueOf(sheetIndex));
		OleAutomation cell3 = WorksheetUtils.getRangeAutomation(annotationDataSheet, ((char) (startColumn+2))+""+endRow, null);
		RangeUtils.setValue(cell3, label);
		OleAutomation cell4 = WorksheetUtils.getRangeAutomation(annotationDataSheet, ((char) (startColumn+3))+""+endRow, null);
		RangeUtils.setValue(cell4, String.valueOf(index));
		OleAutomation cell5 = WorksheetUtils.getRangeAutomation(annotationDataSheet, ((char) (startColumn+4))+""+endRow, null);
		RangeUtils.setValue(cell5, range);
		
		WorksheetUtils.protectWorksheet(annotationDataSheet);	
	}
	
	
	public static OleAutomation createAnnotationDataSheet(OleAutomation workbookAutomation){
		
		WorkbookUtils.unprotectWorkbook(workbookAutomation);
		
		OleAutomation newWorksheet = WorkbookUtils.addWorksheetAsLast(workbookAutomation);
		WorksheetUtils.setWorksheetName(newWorksheet, AnnotationDataSheetName);
	
		OleAutomation cellA1 = WorksheetUtils.getRangeAutomation(newWorksheet, "A1", null);
		RangeUtils.setValue(cellA1, "SheetName");
		OleAutomation cellB1 = WorksheetUtils.getRangeAutomation(newWorksheet, "B1", null);
		RangeUtils.setValue(cellB1, "SheetIndex");
		OleAutomation cellC1 = WorksheetUtils.getRangeAutomation(newWorksheet, "C1", null);
		RangeUtils.setValue(cellC1, "AnnotationLabel");
		OleAutomation cellD1 = WorksheetUtils.getRangeAutomation(newWorksheet, "D1", null);
		RangeUtils.setValue(cellD1, "AnnotationIndex");
		OleAutomation cellE1 = WorksheetUtils.getRangeAutomation(newWorksheet, "E1", null);
		RangeUtils.setValue(cellE1, "Range");
		
		//OleAutomation cellF1 = WorksheetUtils.getRangeAutomation(newWorksheet, "F1", null);
		//RangeUtils.setValue(cellF1, "IsMerged");
		
		//WorksheetUtils.setWorksheetVisibility(newWorksheet, false);		
		WorkbookUtils.protectWorkbook(workbookAutomation, true, false);
		
		return newWorksheet;
	}
	
	/**
	 * 
	 * @param label
	 * @return
	 */
	public static int getAnnotationIndex(String label){
		
		int index = 1;
		if(annotationMap.get(label)==null){
			annotationMap.put(label,index);
		}else{
			index = annotationMap.get(label) + 1 ;
			annotationMap.put(label, index);
		}	
		
		return index;
	}

	/**
	 * Call the appropriate annotation function for the given annotation class
	 * @param workbookAutomation
	 * @param sheetName
	 * @param selectedAreas
	 * @param annotationClass
	 */
	public static void callAnnotationMethod(OleAutomation workbookAutomation, String sheetName, int sheetIndex,
														String selectedAreas[], AnnotationClass annotationClass ){
		
		 switch (annotationClass.getAnnotationTool()) {
		    case RECTANGLE  : annotateSelectedAreasWithRectangle(workbookAutomation, sheetName, sheetIndex, selectedAreas, annotationClass); break;
		    case TEXTBOX  : annotateSelectedAreasWithTextboxes(workbookAutomation, sheetName, sheetIndex, selectedAreas, annotationClass); break;
		    case BORDERAROUND: annotateByBorderingSelectedAreas(workbookAutomation, sheetName, sheetIndex, selectedAreas, annotationClass); break;
		    default: System.out.println("Option not recognized"); System.exit(1); break;
		}
	
	}
	
	
	public static void clearShapeAnnotationsFromActiveSheet(OleAutomation workbookAutomation, String activeSheetName ){
		
		// get the OleAutomation object for the active worksheet using its name
		OleAutomation worksheetAutomation = WorkbookUtils.getWorksheetAutomationByName(workbookAutomation, activeSheetName);
				
		// unprotect the worksheet in order to create the annotation
		WorksheetUtils.unprotectWorksheet(worksheetAutomation);
		
		// delete all shapes that are used for annotating ranges of cells
		OleAutomation shapesAutomation = WorksheetUtils.getWorksheetShapes(worksheetAutomation);	
		int count = CollectionsUtils.getNumberOfObjectsInOleCollection(shapesAutomation);	
		int processed = 0; 
		// all shapes that are used for annotating have names that start with the following string pattern 
		String pattern = activeSheetName.replace(" ", "_")+"_Annotation_";    
		while (processed!=count){
			 OleAutomation shapeAutomation = CollectionsUtils.getItemByIndex(shapesAutomation, 1, true);	 
			 String name = ShapeUtils.getShapeName(shapeAutomation);
			 if(name.indexOf(pattern)>-1){
				 ShapeUtils.deleteShape(shapeAutomation);
			 }
			 processed++;
		}			
		shapesAutomation.dispose();
		
		// protect the worksheet from user manipulation 
		WorksheetUtils.protectWorksheet(worksheetAutomation);
		worksheetAutomation.dispose();
	}
	
//	public static void clearAnnotationDataForActiveSheet(OleAutomation workbookAutomation, String activeSheetName ){
//		
//		// get the OleAutomation object for the active worksheet using its name
//		OleAutomation annotationDataSheet = WorkbookUtils.getWorksheetAutomationByName(workbookAutomation, AnnotationDataSheetName) ;
//				
//		// unprotect the worksheet in order to create the annotation
//		WorksheetUtils.unprotectWorksheet(annotationDataSheet);
//		
//		// filter those annotation data related to the active worksheet. 
//		OleAutomation usedRange = WorksheetUtils.getUsedRange(annotationDataSheet);		
//		RangeUtils.filterRange(usedRange, 1, activeSheetName);
//		
//		OleAutomation filteredRange = RangeUtils.getSpecialCells(usedRange, 12); // xlCellTypeVisible = 12  (visible cells)
//		String address = RangeUtils.getRangeAddress(filteredRange);
//		System.out.println(address);
//		int strIndex = address.indexOf(":");
//		String topDownCell = address.substring(strIndex+1).replace("$","");
//		System.out.println(topDownCell);
//		filteredRange.dispose();
//		usedRange.dispose();
//		
//		OleAutomation rangeExcludeHeader = WorksheetUtils.getRangeAutomation(annotationDataSheet, "A2", topDownCell);
//		OleAutomation dataToDelete = RangeUtils.getSpecialCells(rangeExcludeHeader, 12); // xlCellTypeVisible = 12  (visible cells)
//		RangeUtils.deleteRange(dataToDelete);
//		rangeExcludeHeader.dispose();
//		dataToDelete.dispose();
//		
//		// protect the worksheet from user manipulation 
//		WorksheetUtils.protectWorksheet(annotationDataSheet);
//		annotationDataSheet.dispose();
//	}
	
	public static boolean exportAnnotationsAsCSV(OleAutomation workbookAutomation, String directoryPath, String fileName){
		
		// get the OleAutomation object for the active worksheet using its name
		OleAutomation annotationDataSheet = WorkbookUtils.getWorksheetAutomationByName(workbookAutomation, AnnotationDataSheetName);
		
		// unprotect the worksheet in order to create the annotation
		WorksheetUtils.unprotectWorksheet(annotationDataSheet);
				
		int index = fileName.lastIndexOf('.');
		if (index != -1) {				
			String nameWithoutExtension = fileName.substring(0,index);
			String annotationDataFile = directoryPath+"\\"+nameWithoutExtension+"_annotation_data";
			return WorksheetUtils.saveAsCSV(annotationDataSheet, annotationDataFile);
		}			
		return false;
	}
}
