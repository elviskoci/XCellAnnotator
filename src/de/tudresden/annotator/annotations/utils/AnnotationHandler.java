/**
 * 
 */
package de.tudresden.annotator.annotations.utils;

import java.util.HashMap;

import org.eclipse.swt.ole.win32.OleAutomation;

import de.tudresden.annotator.annotations.AnnotationClass;
import de.tudresden.annotator.annotations.AnnotationTool;
import de.tudresden.annotator.annotations.RangeAnnotation;
import de.tudresden.annotator.oleutils.CharactersUtils;
import de.tudresden.annotator.oleutils.CollectionsUtils;
import de.tudresden.annotator.oleutils.ColorFormatUtils;
import de.tudresden.annotator.oleutils.FillFormatUtils;
import de.tudresden.annotator.oleutils.FontUtils;
import de.tudresden.annotator.oleutils.LineFormatUtils;
import de.tudresden.annotator.oleutils.RangeUtils;
import de.tudresden.annotator.oleutils.ShadowFormatUtils;
import de.tudresden.annotator.oleutils.ShapeUtils;
import de.tudresden.annotator.oleutils.TextFrameUtils;
import de.tudresden.annotator.oleutils.WorkbookUtils;
import de.tudresden.annotator.oleutils.WorksheetUtils;

/**
 * @author Elvis Koci
 */
public class AnnotationHandler {

	private static HashMap<String, Integer> annotationMap = new HashMap<String, Integer>();
	
	public static void annotate(OleAutomation workbookAutomation, String sheetName, int sheetIndex,
								String selectedAreas[], AnnotationClass annotationClass) {

		// get the OleAutomation object for the worksheet using its name
		OleAutomation sheetAutomation = WorkbookUtils.getWorksheetAutomationByName(workbookAutomation, sheetName);

		// unprotect the worksheet in order to create the annotations
		WorksheetUtils.unprotectWorksheet(sheetAutomation);

		// for each area in the range create an annotation
		for (String area : selectedAreas) {
			
			String[] subStrings = area.split(":");

			String topRightCell = subStrings[0];
			String downLeftCell = null;
			if (subStrings.length == 2)
				downLeftCell = subStrings[1];

			OleAutomation rangeAutomation = WorksheetUtils.getRangeAutomation(sheetAutomation, topRightCell, downLeftCell);	
			
			String annotationName = generateAnnotationName(sheetName, annotationClass.getLabel());
			
			callAnnotationMethod(sheetAutomation, rangeAutomation, annotationClass, annotationName);
			
			// save metadata about the annotation
			RangeAnnotation ra = new RangeAnnotation(sheetName, sheetIndex, annotationClass, annotationName, area);
			AnnotationData.saveAnnotationData(workbookAutomation, ra);	
		}
					
		// protect the worksheet to prevent user from modifying the annotations
		WorksheetUtils.protectWorksheet(sheetAutomation);
		WorksheetUtils.makeWorksheetActive(sheetAutomation);
		sheetAutomation.dispose();
	}
	
	/**
	 * Call an annotation function depending on the annotation tool
	 * @param annotationClass
	 * @param rangeAutomation
	 * @param sheetAutomation
	 * @param annotationName
	 */
	public static void callAnnotationMethod(OleAutomation sheetAutomation, OleAutomation rangeAutomation, 
															AnnotationClass annotationClass, String annotationName){
		switch (annotationClass.getAnnotationTool()) {
		case SHAPE  : annotateSelectedAreaWithShapes(sheetAutomation, rangeAutomation, annotationClass, annotationName); break;
		case TEXTBOX  :  annotateSelectedAreaWithShapes(sheetAutomation, rangeAutomation, annotationClass, annotationName); break;
		case BORDERAROUND: annotateByBorderingSelectedArea(rangeAutomation, annotationClass, annotationName); break;
		default: System.out.println("Option not recognized"); System.exit(1); break;
		}	
	}
	
	
	/**
	 * 
	 * @param rangeAutomation
	 * @param annotationClass
	 * @param annotationName
	 */
	public static void annotateByBorderingSelectedArea(OleAutomation rangeAutomation, AnnotationClass annotationClass, String annotationName){
	
		long color = annotationClass.getLineColor();
		if(annotationClass.getLineColor()<0){
			color = annotationClass.getColor();
		}
			
		RangeUtils.drawBorderAroundRange(rangeAutomation, annotationClass.getLineStyle(), annotationClass.getLineWeight(), color);	
		rangeAutomation.dispose();
	}
	
	/**
	 * 
	 * @param sheetAutomation
	 * @param rangeAutomation
	 * @param annotationClass
	 * @param annotationName
	 */
	public static void annotateSelectedAreaWithShapes(OleAutomation sheetAutomation, OleAutomation rangeAutomation, 
																		AnnotationClass annotationClass, String annotationName){
		
		double left = RangeUtils.getRangeLeftPosition(rangeAutomation);  
		double top = RangeUtils.getRangeTopPosition(rangeAutomation);
		double width = RangeUtils.getRangeWidth(rangeAutomation);
		double height = RangeUtils.getRangeHeight(rangeAutomation);
		rangeAutomation.dispose();
		
		OleAutomation shapesAutomation = WorksheetUtils.getWorksheetShapes(sheetAutomation);	
		
		if(annotationClass.getAnnotationTool()==AnnotationTool.TEXTBOX){	
			
			OleAutomation textboxAutomation = ShapeUtils.drawTextBox(shapesAutomation, left, top, width, height); 
			setAnnotationProperties(textboxAutomation, annotationClass);
			ShapeUtils.setShapeName(textboxAutomation, annotationName);
			textboxAutomation.dispose();
		}
		
		if(annotationClass.getAnnotationTool()==AnnotationTool.SHAPE){		
			
			OleAutomation shapeAuto = ShapeUtils.drawShape(shapesAutomation, annotationClass.getShapeType(), left, top, width, height);
			setAnnotationProperties(shapeAuto, annotationClass);  
			ShapeUtils.setShapeName(shapeAuto, annotationName);
			shapeAuto.dispose();
		}
				
		shapesAutomation.dispose();
	}
			
	
	/**
	 * Format the annotation object (shape, textbox, etc) used to annotate 
	 * @param annotation an OleAutomation to access the functionalities of the annotation object
	 * @param annotationClass an object that represents an annotation class
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
	
		// set text properties
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
	 * Hide/Show all shapes in the embedded workbook that are used for annotating ranges of cells
	 * @param workbookAutomation an OleAutomation to access the functionalities of the embedded workbook
	 * @param visible true to make annotation shapes visible, false to hide them
	 */
	public static void setVisilityForShapeAnnotations(OleAutomation workbookAutomation, boolean visible){
		
		OleAutomation worksheets = WorkbookUtils.getWorksheetsAutomation(workbookAutomation);
		int count = CollectionsUtils.countItemsInCollection(worksheets);
		
		// indexing starts from 1 
		for (int i = 1; i < count; i++) {
			OleAutomation sheet = CollectionsUtils.getItemByIndex(worksheets, i, false);
			setVisibilityForWorksheetShapeAnnotations(sheet, visible);
		}		
	}
	
	
	/**
	 * Hide/Show all the shapes used for annotations in the worksheet having the given name. 
	 * This method will loop through the collection of shapes in the worksheet and set their visibility to false or true,
	 * if they are used for annotations. The other shapes, that were present in the original file, will not be affected. 
	 * @param workbookAutomation an OleAutomation to access the functionalities of the embedded workbook
	 * @param sheetName the name of the worksheet for which the action will be performed
	 * @param visible true to make annotation shapes visible, false to hide them
	 */
	public static void setVisibilityForWorksheetShapeAnnotations(OleAutomation workbookAutomation, String sheetName, boolean visible ){
		
		// worksheet that stores the annotation (meta-)data does not have annotation shapes 
		if(sheetName.compareTo(AnnotationData.name)==0)
			return;
		
		// get the OleAutomation object for the worksheet using the given name
		OleAutomation worksheetAutomation = WorkbookUtils.getWorksheetAutomationByName(workbookAutomation, sheetName);
		setVisibilityForWorksheetShapeAnnotations(worksheetAutomation, visible);
	}
	
	
	/**
	 * Hide/Show all the shapes used for annotations in the worksheet having the given name. 
	 * This method will loop through the collection of shapes in the worksheet and set their visibility to false or true,
	 * if they are used for annotations. The other shapes, that were present in the original file, will not be affected.
	 * @param worksheetAutomation an OleAutoamtion to access the functionalities of the worksheet that action will be applied on
	 * @param visible true to make annotation shapes visible, false to hide them
	 */
	public static void setVisibilityForWorksheetShapeAnnotations(OleAutomation worksheetAutomation,  boolean visible){
		
		// unprotect the worksheet in order to change the visibility of the shapes
		WorksheetUtils.unprotectWorksheet(worksheetAutomation);
		
		// get the collection of shapes in the worksheet
		OleAutomation shapesAutomation = WorksheetUtils.getWorksheetShapes(worksheetAutomation);
		
		// if there are no shapes in the worksheet, skip the process
		if(shapesAutomation==null)
			return;
		 		
		// all shapes that are used for annotations have names that start with the following string pattern 
		String sheetName = WorksheetUtils.getWorksheetName(worksheetAutomation);
		String startString = getStartOfAnnotationName(sheetName);
		
		int count = CollectionsUtils.countItemsInCollection(shapesAutomation);	
		for (int i = 1; i <= count; i++) {
			 OleAutomation shapeAutomation = CollectionsUtils.getItemByIndex(shapesAutomation, i, true);	 
			 String name = ShapeUtils.getShapeName(shapeAutomation);
			 if(name.indexOf(startString)== 0){
				 ShapeUtils.setShapeVisibility(shapeAutomation, visible);
			 }
			 shapeAutomation.dispose();
		}
				
		shapesAutomation.dispose();
		
		// protect the worksheet from further user manipulation 
		WorksheetUtils.protectWorksheet(worksheetAutomation);
		worksheetAutomation.dispose();	
	}
	
	
	/**
	 * Delete all shapes in the embedded workbook that are used for annotating ranges of cells
	 * @param workbookAutomation an OleAutomation to access the functionalities of the embedded workbook
	 */
	public static void deleteAllShapeAnnotations(OleAutomation workbookAutomation){
		
		OleAutomation worksheets = WorkbookUtils.getWorksheetsAutomation(workbookAutomation);		
		int count = CollectionsUtils.countItemsInCollection(worksheets);
		
		// indexing starts from 1 
		for (int i = 1; i < count; i++) {
			OleAutomation sheet = CollectionsUtils.getItemByIndex(worksheets, i, false);
			deleteShapeAnnotationsFromWorksheet(sheet);
		}
	}
	
	
	/**
	 * Delete all the shapes used for annotations in the worksheet having the given name 
	 * @param workbookAutomation an OleAutomation to access the functionalities of the embedded workbook
	 * @param sheetName the name of the worksheet for which the annotation shapes will be deleted
	 */
	public static void deleteShapeAnnotationsFromWorksheet(OleAutomation workbookAutomation, String sheetName ){
		
		// can not apply this method to the worksheet that stores the annotation (meta-)data, as it does not contain any shapes 
		if(sheetName.compareTo(AnnotationData.name)==0)
			return;
		
		// get the OleAutomation object for the active worksheet using its name
		OleAutomation worksheetAutomation = WorkbookUtils.getWorksheetAutomationByName(workbookAutomation, sheetName);
		deleteShapeAnnotationsFromWorksheet(worksheetAutomation);
	}

	
	/**
	 * Delete all the shapes used for annotations in the given worksheet 
	 * @param worksheetAutomation an OleAutoamtion to access the functionalities of the worksheet that action will be applied on
	 */
	public static void deleteShapeAnnotationsFromWorksheet(OleAutomation worksheetAutomation){
		
		// unprotect the worksheet in order delete the shape annotations
		WorksheetUtils.unprotectWorksheet(worksheetAutomation);
		
		// delete all shapes that are used for annotating ranges of cells
		OleAutomation shapesAutomation = WorksheetUtils.getWorksheetShapes(worksheetAutomation);	
		int count = CollectionsUtils.countItemsInCollection(shapesAutomation);	
		int processed = 0; 
		
		// all shapes that are used for annotating have names that start with the following string pattern 
		String sheetName = WorksheetUtils.getWorksheetName(worksheetAutomation);
		String startString =  getStartOfAnnotationName(sheetName);
		while (processed!=count){
			 OleAutomation shapeAutomation = CollectionsUtils.getItemByIndex(shapesAutomation, 1, true);	 
			 String name = ShapeUtils.getShapeName(shapeAutomation);
			 if(name.indexOf(startString)==0){
				 ShapeUtils.deleteShape(shapeAutomation);
			 }
			 processed++;
			 // TODO: Dispose shapeAutomation ?
		}			
		shapesAutomation.dispose();
		
		// protect the worksheet from further user manipulation 
		WorksheetUtils.protectWorksheet(worksheetAutomation);
		worksheetAutomation.dispose();	
	}
	
	
	/**
	 * Get the index of the annotation. This is a value that is incremented each time a new annotation is created.
	 * @param label the Annotation class label
	 * @return an integer that represents the index of the annotation 
	 */
	public static int generateAnnotationIndex(String label){
		
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
	 * Get the name of the annotation 
	 * @param sheetName the name of the worksheet that contains the annotation
	 * @param classLabel the label of the class the annotation is member of
	 * @return a string that represents the annotation name
	 */
	public static String generateAnnotationName(String sheetName, String classLabel){	
		 int annotationIndex = generateAnnotationIndex(classLabel);
		 String startOfAnnotationName = getStartOfAnnotationName(sheetName);		 
		 String formatedName = startOfAnnotationName+"_"+classLabel+"_"+annotationIndex;
		 
		 return formatedName.toUpperCase();
	}
	
	
	/**
	 * Get the string pattern that it is contained in the name of all annotation objects of a worksheet.   
	 * @param sheetName 
	 * @return
	 */
	public static String getStartOfAnnotationName(String sheetName){
		
		String name = sheetName.replace(" ", "_")+"_Annotation"; 
		return  name.toUpperCase(); 
	}
}
