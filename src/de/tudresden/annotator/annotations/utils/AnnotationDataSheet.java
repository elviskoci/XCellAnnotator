/**
 * 
 */
package de.tudresden.annotator.annotations.utils;

import java.util.HashMap;
import java.util.Iterator;
import java.util.LinkedHashMap;

import org.eclipse.swt.ole.win32.OleAutomation;

import de.tudresden.annotator.annotations.AnnotationClass;
import de.tudresden.annotator.annotations.RangeAnnotation;
import de.tudresden.annotator.annotations.WorkbookAnnotation;
import de.tudresden.annotator.annotations.WorksheetAnnotation;
import de.tudresden.annotator.oleutils.CollectionsUtils;
import de.tudresden.annotator.oleutils.RangeUtils;
import de.tudresden.annotator.oleutils.WorkbookUtils;
import de.tudresden.annotator.oleutils.WorksheetUtils;

/**
 * @author Elvis Koci
 */
public class AnnotationDataSheet {
	
	protected static final String name = "Annotation_Data_Sheet";
	private static String startColumn = "A";
	private static int startRow = 1; 
	
	/**
	 * This linked hash map stores the names of the fields used in the Annotation Data Sheet and their default order. 
	 */
	private static final LinkedHashMap<String, Integer> fields;
	static
    {
		fields = new LinkedHashMap<String,Integer>();
		fields.put("Sheet.Name", 0); // required 
		fields.put("Sheet.Index", 1); // required
		fields.put("Annotation.Label", 2); // required
		fields.put("Annotation.Name", 3); // required
		// fields.put("AnnotationTool.Name", 4); // required
		// fields.put("AnnotationTool.Code", 5); // required  
		fields.put("Annotation.Range", 4); // required
		fields.put("Annotation.Parent", 5); // required
    }
	

	/**
	 * Save new annotation data
	 * @param workbookAutomation an OleAutomation to access the embedded workbook
	 * @param annotation a RangeAnnotation object that maintains (contains) the annotation data to be saved
	 */
	public static void saveAnnotationData(OleAutomation workbookAutomation, RangeAnnotation annotation){
			
		OleAutomation annotationDataSheet =  WorkbookUtils.getWorksheetAutomationByName(workbookAutomation, name);
		
		if(annotationDataSheet==null){		
			annotationDataSheet = createAnnotationDataSheet(workbookAutomation);			
		}
		
		OleAutomation usedRange = WorksheetUtils.getUsedRange(annotationDataSheet);		
		String usedAddress = RangeUtils.getRangeAddress(usedRange);
		usedRange.dispose();
		
		usedAddress = usedAddress.replace("$", "");
		String[] cells =  usedAddress.split(":");
				
		int endRow =   Integer.valueOf(cells[1].replaceAll("[^0-9]+",""));
		int row =  endRow + 1;
					
		writeNewRow( annotationDataSheet, row, annotation);		
		annotationDataSheet.dispose();
	}
	
	
	/**
	 * Create the sheet that will store the annotation data 
	 * @param workbookAutomation an OleAutomation to access the embedded workbook
	 * @return the OleAutomation of the created worksheet
	 */
	protected static OleAutomation createAnnotationDataSheet(OleAutomation workbookAutomation){
		
		WorkbookUtils.unprotectWorkbook(workbookAutomation);
		
		OleAutomation newWorksheet = WorkbookUtils.addWorksheetAsLast(workbookAutomation);
		WorksheetUtils.setWorksheetName(newWorksheet, name);
		
		char lastChar = startColumn.charAt(startColumn.length() - 1 ); 
		String startChars = startColumn.substring(0, startColumn.length() - 1);
	
		int row = startRow; 
	
		Iterator<String> itr = fields.keySet().iterator();
		int i = 0;	
		while (itr.hasNext()) {
			String cellAddress = "$"+(startChars+((char) (lastChar+i))+"$"+(row));
			OleAutomation cell = WorksheetUtils.getRangeAutomation(newWorksheet, cellAddress, null);
			RangeUtils.setValue(cell, itr.next());
			i++;
		}
				
		//WorksheetUtils.setWorksheetVisibility(newWorksheet, false);		
		WorkbookUtils.protectWorkbook(workbookAutomation, true, false);
		
		return newWorksheet;
	}
	
	
	/**
	 * Write new row of annotation data
	 * @param annotationDataSheet an OleAutomation that provides access to the sheet that maintains the annotation data
	 * @param row an integer that represents the index of the row to write the data
	 * @param annotation a RangeAutomation object that maintains (contains) the annotation data to write  
	 */
	protected static void writeNewRow(OleAutomation annotationDataSheet, int row, RangeAnnotation annotation){		
		
		WorksheetUtils.unprotectWorksheet(annotationDataSheet);
		
		//TODO: get next (adjacent) column using OleAutomation 
		char lastChar = startColumn.charAt(startColumn.length() - 1 ); 
		String startChars = startColumn.substring(0, startColumn.length() - 1);
		
		Iterator<String> itr = fields.keySet().iterator();
		int i = 0;	
		while (itr.hasNext()) {
			OleAutomation cell = WorksheetUtils.getRangeAutomation(annotationDataSheet, startChars+""+((char) (lastChar+i))+""+row, null);
			RangeUtils.setValue(cell, getFieldData(itr.next(), annotation));	
			i++;
		}
		
		WorksheetUtils.protectWorksheet(annotationDataSheet);	
	}
	
	
	/**
	 * Get the value for the field from the corresponding attribute/s of the annotation object 
	 * @param fieldName a string that represents the name of a field from the header row in the annotation data sheet 
	 * @param annotation a RangeAutomation object that maintains (contains) the annotation data to be retrieved  
	 * @return a string that represents the value of the specified (given) field
	 */
	public static String getFieldData(String fieldName, RangeAnnotation annotation){
		
		String value = null;		
		switch (fieldName) {
			case "Sheet.Name"  : value = annotation.getSheetName(); break;
			case "Sheet.Index"  : value = String.valueOf(annotation.getSheetIndex()); break;
			case "Annotation.Label"  : value = annotation.getAnnotationClass().getLabel(); break;
			case "Annotation.Name"  : value = annotation.getName(); break;
			case "AnnotationTool.Name"  : value = annotation.getAnnotationClass().getAnnotationTool().name(); break;
			case "AnnotationTool.Code"  : value = String.valueOf(annotation.getAnnotationClass().getAnnotationTool().getCode()); break;
			case "Annotation.Range"  : value = annotation.getRangeAddress(); break;
			case "Annotation.Parent" : value = annotation.getParent() instanceof  RangeAnnotation ? 
											   ((RangeAnnotation) annotation.getParent()).getName() : annotation.getSheetName(); break;
			default: value = "Field not recognized!"; break;
		}
		return value;
	}
	
		
	/**
	 * Read all annotation data from the "Annotation Data" Sheet 
	 * @param workbookAutomation an OleAutomation that provides access to the embedded workbook
	 * @return true if annotation data were successfully read, false otherwise
	 */
	public static boolean readAnnotationData(OleAutomation workbookAutomation){
		
		// get the OleAutomation object for the sheet that stores 
		// the annotation metadata (a.k.a. annotation data sheet) 
		OleAutomation annotationDataSheet = 
				WorkbookUtils.getWorksheetAutomationByName(workbookAutomation, name);
		if(annotationDataSheet==null){
			System.out.println("No data to read. Annotation data sheet not found!");
			return false;
		}
		
		// get the range that has data 
		OleAutomation usedRange = WorksheetUtils.getUsedRange(annotationDataSheet);
		if(usedRange==null){
			System.out.println("No data to read. All cells are empty (not used)!");
			return false;
		}
		
		// get used areas 
		OleAutomation usedAreas = RangeUtils.getAreas(usedRange);
		int countAreas = CollectionsUtils.countItemsInCollection(usedAreas);
		usedAreas.dispose();
		
		if(countAreas>1){
			System.err.println("Could not read the annotation data from the sheet "+name+".\n"+
								"Data are not in the expected format!");
			System.exit(1);
		}
		
		if(countAreas==0){
			System.out.println("No data to read. All cells are empty (not used)!");
			return false;
		}
		
		String usedRangeAddress = RangeUtils.getRangeAddress(usedRange);
		usedRange.dispose();
		
		String boundingCells[] = usedRangeAddress.split(":"); 
		String topLeftCell = boundingCells[0];
		String topLeftColumn = topLeftCell.replaceAll("[0-9\\$]+","");
		int topLeftRow = Integer.valueOf(topLeftCell.replaceAll("[^0-9]+",""));
		String downRightCell = boundingCells[1];
		String downRightColumn = downRightCell.replaceAll("[0-9\\$]+","");
		int downRightRow = Integer.valueOf(downRightCell.replaceAll("[^0-9]+",""));
		
		boolean  result = validateHeaderRow(annotationDataSheet, topLeftRow, topLeftColumn, downRightColumn);
		if(!result){
			System.exit(1);
		}
		
		WorkbookAnnotation workbookAnnotation = AnnotationHandler.getWorkbookAnnotation();
		for (int i = (topLeftRow + 1); i <=downRightRow; i++) {
			RangeAnnotation annotation = readAnnotationDataRow(annotationDataSheet, i, topLeftColumn, downRightColumn);
			workbookAnnotation.addRangeAnnotation(annotation);
		}
		
		return true;
	}
	
	
	/**
	 * Validate header row. It should contain all the expected (predefined) fields  
	 * @param annotationDataSheet an OleAutomation to access the functionalities of the worksheet that stores the annotation data
	 * @param topLeftRow an integer that represents the address of the top left row
	 * @param topLeftColumn a string that represents the address of the column on the top left
	 * @param downRightColumn a string that represents the address of the column on the down right
	 * @return true if the header row passes all checks, false if validation fails. 
	 */
	protected static boolean validateHeaderRow(OleAutomation annotationDataSheet, int topLeftRow, String topLeftColumn, String downRightColumn){
		
		// get all the values from the range that represents the header
		String topLeftCell = topLeftColumn+topLeftRow;
		String downRightCell = downRightColumn+topLeftRow;		
		OleAutomation rangeAutomation = WorksheetUtils.getRangeAutomation(annotationDataSheet, topLeftCell, downRightCell);
		String values[] = RangeUtils.getRangeValues(rangeAutomation);
		rangeAutomation.dispose();
		
		// check if the number of fields in the sheet match with the pre-defined (expected) ones.  
		if(values.length < fields.size()){
			System.out.println("The number of fields in the annotation data sheet is larger than the declared (expected) fields");
			return false;
		}else if(values.length > fields.size()){
			System.out.println("The number of fields in the annotation data sheet is smaller than the declared (expected) fields");
			return false;
		}

		// check that the header row contains recognizable fields. update their order
		for (int i = 0; i< values.length; i++) {
			
			String val = values[i];
			if( val==null || val.compareTo("")==0){
				System.err.println("Empty cells in the header row!");
				return false;
			}
			
			if(!fields.containsKey(val)){
				System.err.println("Field not recognized");
				return false;
			}
			
			fields.put(val, i);
		} 
		
		// update the start column and row (e.i., the address of the first cell) of the range that contains the annotation data
		startColumn = topLeftColumn;
		startRow = topLeftRow;
				
		return true;
	}
	
	
	/**
	 * Read a row of annotation data and create re-create the RangeAnnotation object
	 * @param annotationDataSheet an OleAutomation to access the functionalities of the worksheet that stores the annotation data
	 * @param row an integer that represents the row to be read
	 * @param startColumn a string that represents the column to start the reading
	 * @param endColumn a string that represents the column to end the reading
	 * @return a RangeAnnotation object that is created using the data read from the row
	 */
	protected static RangeAnnotation readAnnotationDataRow(OleAutomation annotationDataSheet, int row, String startColumn, String endColumn){		
		
		// get the values of the cells in the row, bounded by the start and the end column.   
		String topLeftCell = startColumn+row;
		String downRightCell = endColumn+row;		
		OleAutomation rangeAutomation = WorksheetUtils.getRangeAutomation(annotationDataSheet, topLeftCell, downRightCell);
		String values[] = RangeUtils.getRangeValues(rangeAutomation);
		rangeAutomation.dispose();
		
		// create a RangeAnnotation object from the annotation data in the row.
		String sheetName = values[fields.get("Sheet.Name")];
		int sheetIndex = Integer.valueOf(values[fields.get("Sheet.Index")]);
		String annotationLabel = values[fields.get("Annotation.Label")];
		String annotationName = values[fields.get("Annotation.Name")];
		String rangeAddress =  values[ fields.get("Annotation.Range")];
		String parent = values[ fields.get("Annotation.Parent")];				
		
		AnnotationClass annotationClass = ClassGenerator.getAnnotationClasses().get(annotationLabel);		
		RangeAnnotation annotation = new RangeAnnotation(sheetName, sheetIndex, annotationClass, annotationName, rangeAddress); 
		
		// set the parent of the RangeAutomation object 
		WorkbookAnnotation workbookAnnotation = AnnotationHandler.getWorkbookAnnotation();
		if(!annotationClass.isContainable()){	
			WorksheetAnnotation worksheetAnnotation = workbookAnnotation.getWorksheetAnnotations().get(parent);
			if(worksheetAnnotation == null){
				workbookAnnotation.getWorksheetAnnotations().put(sheetName, new WorksheetAnnotation(sheetName, sheetIndex));
			}
			annotation.setParent(worksheetAnnotation);
		}else{		
			if(annotationClass.isDependent()){
				AnnotationClass parentClass = annotationClass.getContainer();
				for (RangeAnnotation ra : workbookAnnotation.getSheetAnnotationsByClass(sheetName, parentClass.getLabel())){
					if(ra.getName().compareTo(parent)==0){
						annotation.setParent(ra);
						break;
					}
				}
			}else{
				WorksheetAnnotation worksheetAnnotation = workbookAnnotation.getWorksheetAnnotations().get(parent);
				if(worksheetAnnotation == null){
					Iterator<AnnotationClass> itr = ClassGenerator.getAnnotationClasses().values().iterator();
					while (itr.hasNext()) {
						AnnotationClass ac = (AnnotationClass) itr.next();
						if(ac.isContainer()){
							for (RangeAnnotation ra : workbookAnnotation.getSheetAnnotationsByClass(sheetName, ac.getLabel())){
								if(ra.getName().compareTo(parent)==0){
									annotation.setParent(ra);
									break;
								}
							}								
						}
					}
				}else{
					annotation.setParent(worksheetAnnotation);	
				}
			}
		}		
		return annotation;
	}
	
	
	/**
	 * Clear the annotation data for the worksheet with the given name 
	 * This method will clear all the rows in the annotation (meta-)data sheet 
	 * that are associated with the specified worksheet  
	 * @param workbookAutomation
	 * @param sheetName
	 */
	public static void deleteAnnotationDataForWorksheet(OleAutomation workbookAutomation, String sheetName ){
		
		// the sheet that stores the annotation metadata is excluded from this process
		if(sheetName.compareTo(name)==0)
			return;
			
		// get the OleAutomation object for the sheet that stores the annotation metadata (a.k.a. annotation data sheet) 
		OleAutomation annotationDataSheetBeforeDelete = WorkbookUtils.getWorksheetAutomationByName(workbookAutomation, name);
		
		// unprotect the annotation data sheet
		boolean isUnprotected= WorksheetUtils.unprotectWorksheet(annotationDataSheetBeforeDelete);
		if(!isUnprotected){
			System.out.println("Annotation Data Sheet could not be unprotected!");
			System.exit(1);
		}
				
		// filter those annotation data rows that contain data about the given worksheet name
		OleAutomation usedRange = WorksheetUtils.getUsedRange(annotationDataSheetBeforeDelete);		
		RangeUtils.filterRange(usedRange, 1, sheetName); // 1 corresponds to the column of worksheet names 
		
		// get the range OleAutomation for the filtered results. This range includes the header row.  
		OleAutomation filteredRange = RangeUtils.getSpecialCells(usedRange, 12); // xlCellTypeVisible = 12  (visible cells)
		OleAutomation areas = RangeUtils.getAreas(filteredRange);
		usedRange.dispose();
		filteredRange.dispose();
		
		// delete all filtered rows except of the header 		
		int count = CollectionsUtils.countItemsInCollection(areas);
		int processed = 0; 
		int i = 1; 	
		while(processed!=count){					
			OleAutomation range = CollectionsUtils.getItemByIndex(areas, i, false); 
			String address = RangeUtils.getRangeAddress(range);
			String topLeftCell = address.substring(0, 4);
			String downRightCell = address.substring(5);
			int position =  downRightCell.lastIndexOf("$");
			int lastRow  = Integer.valueOf(downRightCell.substring(position+1));
			
			if(topLeftCell.compareTo("$A$1")==0){
				if(lastRow==1){
					i = 2;
				}else{
					RangeUtils.deleteRange(WorksheetUtils.getRangeAutomation( annotationDataSheetBeforeDelete,"$A$2",downRightCell));
				}
			}else {
				RangeUtils.deleteRange(range);
			}
			processed++;
		}		
		areas.dispose();
		
		annotationDataSheetBeforeDelete.dispose();
		
		// remove filter from the annotation data sheet, to show all data
		OleAutomation annotationDataSheetAfterDelete = WorkbookUtils.getWorksheetAutomationByName(workbookAutomation, name);
		WorksheetUtils.showAllWorksheetData(annotationDataSheetAfterDelete);
		
		// protect the worksheet from further user manipulation 
		WorksheetUtils.protectWorksheet(annotationDataSheetAfterDelete);
		annotationDataSheetAfterDelete.dispose();		
	}
	
	
	/**
	 * Clear all annotation data. This method clears all the values from the sheet 
	 * where the annotation data are stored, except of the header row. 
	 * @param workbookAutomation
	 */
	public static void deleteAllAnnotationData(OleAutomation workbookAutomation){
		
		// get the OleAutomation object for the sheet that stores the annotation metadata (a.k.a. annotation data sheet) 
		OleAutomation annotationDataSheetBeforeDelete = WorkbookUtils.getWorksheetAutomationByName(workbookAutomation, name);
		
		// unprotect the worksheet in order to perform the following actions
		boolean isUnprotected= WorksheetUtils.unprotectWorksheet(annotationDataSheetBeforeDelete);
		if(!isUnprotected){
			System.out.println("Annotation Data Sheet could not be unprotected!");
			System.exit(1);
		}
		
		// find the last row that contains data
		OleAutomation usedRange = WorksheetUtils.getUsedRange(annotationDataSheetBeforeDelete);
		String address = RangeUtils.getRangeAddress(usedRange);
		int position = address.indexOf(":");
		String downRightCell = address.substring(position+1).replace("$", "");
		usedRange.dispose();
		
		// delete all the rows except of the one header
		OleAutomation rangeToDelete = WorksheetUtils.getRangeAutomation(annotationDataSheetBeforeDelete, "A2", downRightCell);
		RangeUtils.deleteRange(rangeToDelete);
		rangeToDelete.dispose();
		
		annotationDataSheetBeforeDelete.dispose();
		
		// protect the worksheet from further user manipulation 
		OleAutomation annotationDataSheetAfterDelete = WorkbookUtils.getWorksheetAutomationByName(workbookAutomation, name);
		WorksheetUtils.protectWorksheet(annotationDataSheetAfterDelete);
		annotationDataSheetAfterDelete.dispose();		
	}
	
	
	/**
	 * Export annotation metadata as CSV
	 * @param workbookAutomation
	 * @param directoryPath
	 * @param fileName
	 * @return
	 */
	public static boolean exportAnnotationsAsCSV(OleAutomation workbookAutomation, String directoryPath, String fileName){
		
		// get the OleAutomation object for the worksheet where the annotation data are stored
		OleAutomation annotationDataSheet = WorkbookUtils.getWorksheetAutomationByName(workbookAutomation, name);
		
		// check if the annotation data sheet exists
		if(annotationDataSheet==null){		
			System.out.println("Annotation Data Sheet not found!");
			return false;
		}
		
		// check if there are annotation data 
		OleAutomation usedRange = WorksheetUtils.getUsedRange(annotationDataSheet);
		if(usedRange==null){
			System.out.println("Used range is null!");
			return false;
		}
		
		// there should be more than two rows of annotation data, 
		// otherwise will export only the header
		OleAutomation rows = RangeUtils.getRangeRows(usedRange);
		if(CollectionsUtils.countItemsInCollection(rows)<2) {
			System.out.println("There are no annotation data, just the header!");
			return false;
		}
		usedRange.dispose();
		rows.dispose();
		
		// unprotect worksheet in order allow export
		WorksheetUtils.unprotectWorksheet(annotationDataSheet);
		
		// export annotation data
		int index = fileName.lastIndexOf('.');		
		String nameWithoutExtension = fileName.substring(0, index);
		String annotationDataFile = directoryPath+"\\"+nameWithoutExtension+"_annotation_data";
		boolean isSuccess = WorksheetUtils.saveAsCSV(annotationDataSheet, annotationDataFile);
		
		// protect worksheet from further user manipulation
		WorksheetUtils.protectWorksheet(annotationDataSheet);
		
		return isSuccess;
	}
	
	
	/**
	 * Hide/Show the worksheet that stores the annotation (meta-)data
	 * @param embeddedWorkbook an OleAutomation that is used to access the functionalities of the workbook that is currently embedded by the application
	 * @param visible true to show the worksheet, false to hide it
	 * @return true if the operation was successful, false otherwise
	 */
	public static boolean setVisibility(OleAutomation embeddedWorkbook, boolean visible){
		
		OleAutomation annotationDataSheet = WorkbookUtils.getWorksheetAutomationByName(embeddedWorkbook, name);
		
		if(annotationDataSheet==null)
			return false; 
		
		boolean result = WorksheetUtils.setWorksheetVisibility(annotationDataSheet, visible);
		annotationDataSheet.dispose();
		return result;
	}
}
