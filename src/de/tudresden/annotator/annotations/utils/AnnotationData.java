/**
 * 
 */
package de.tudresden.annotator.annotations.utils;

import org.eclipse.swt.ole.win32.OleAutomation;

import de.tudresden.annotator.annotations.RangeAnnotation;
import de.tudresden.annotator.annotations.WorkbookAnnotation;
import de.tudresden.annotator.oleutils.CollectionsUtils;
import de.tudresden.annotator.oleutils.RangeUtils;
import de.tudresden.annotator.oleutils.WorkbookUtils;
import de.tudresden.annotator.oleutils.WorksheetUtils;

/**
 * @author Elvis Koci
 */
public class AnnotationData {
	
	protected static final String name = "Annotation_Data_Sheet";
	protected static final WorkbookAnnotation workbookAnnotation = new WorkbookAnnotation();
	
	/**
	 * Save the annotation data
	 * 
	 * @param workbookAutomation
	 * @param annotation
	 */
	public static void saveAnnotationData(OleAutomation workbookAutomation, RangeAnnotation annotation){
			
		OleAutomation annotationDataSheet =  WorkbookUtils.getWorksheetAutomationByName(workbookAutomation, name);
		
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
					
		writeNewRow( annotationDataSheet, startColumn, endRow, annotation);			
		
		
		if(workbookAnnotation.getWorkbookName() == null || workbookAnnotation.getWorkbookName().compareTo("")==0){		
			String workbookName = WorkbookUtils.getWorkbookName(workbookAutomation);
			workbookAnnotation.setWorkbookName(workbookName);
		}
		workbookAnnotation.addRangeAnnotation(annotation);
	}
	
	
	/**
	 * Write new row of annotation data
	 * 
	 * @param annotationDataSheet
	 * @param startColumn
	 * @param endRow
	 * @param annotation
	 */
	protected static void writeNewRow(OleAutomation annotationDataSheet, char startColumn, int endRow, RangeAnnotation annotation){		
		
		WorksheetUtils.unprotectWorksheet(annotationDataSheet);
		
		OleAutomation cell1 = WorksheetUtils.getRangeAutomation(annotationDataSheet, startColumn+""+endRow, null);
		RangeUtils.setValue(cell1, annotation.getSheetName());	 
		OleAutomation cell2 = WorksheetUtils.getRangeAutomation(annotationDataSheet, ((char) (startColumn+1))+""+endRow, null);
		RangeUtils.setValue(cell2, String.valueOf(annotation.getSheetIndex()));
		OleAutomation cell3 = WorksheetUtils.getRangeAutomation(annotationDataSheet, ((char) (startColumn+2))+""+endRow, null);
		RangeUtils.setValue(cell3, annotation.getAnnotationClass().getLabel());
		OleAutomation cell4 = WorksheetUtils.getRangeAutomation(annotationDataSheet, ((char) (startColumn+3))+""+endRow, null);
		RangeUtils.setValue(cell4, annotation.getName());
		OleAutomation cell5 = WorksheetUtils.getRangeAutomation(annotationDataSheet, ((char) (startColumn+4))+""+endRow, null);
		RangeUtils.setValue(cell5, annotation.getAnnotationClass().getAnnotationTool().name());
		OleAutomation cell6 = WorksheetUtils.getRangeAutomation(annotationDataSheet, ((char) (startColumn+5))+""+endRow, null);
		RangeUtils.setValue(cell6, String.valueOf(annotation.getAnnotationClass().getAnnotationTool().getCode()));
		OleAutomation cell7 = WorksheetUtils.getRangeAutomation(annotationDataSheet, ((char) (startColumn+6))+""+endRow, null);
		RangeUtils.setValue(cell7, annotation.getRangeAddress());
		
		WorksheetUtils.protectWorksheet(annotationDataSheet);	
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
	
		OleAutomation cellA1 = WorksheetUtils.getRangeAutomation(newWorksheet, "A1", null);
		RangeUtils.setValue(cellA1, "SheetName");
		OleAutomation cellB1 = WorksheetUtils.getRangeAutomation(newWorksheet, "B1", null);
		RangeUtils.setValue(cellB1, "SheetIndex");
		OleAutomation cellC1 = WorksheetUtils.getRangeAutomation(newWorksheet, "C1", null);
		RangeUtils.setValue(cellC1, "AnnotationLabel");
		OleAutomation cellD1 = WorksheetUtils.getRangeAutomation(newWorksheet, "D1", null);
		RangeUtils.setValue(cellD1, "AnnotationName");
		OleAutomation cellE1 = WorksheetUtils.getRangeAutomation(newWorksheet, "E1", null);
		RangeUtils.setValue(cellE1, "AnnotationToolName");
		OleAutomation cellF1 = WorksheetUtils.getRangeAutomation(newWorksheet, "F1", null);
		RangeUtils.setValue(cellF1, "AnnotationToolCode");
		OleAutomation cellG1 = WorksheetUtils.getRangeAutomation(newWorksheet, "G1", null);
		RangeUtils.setValue(cellG1, "Range");
		
		//WorksheetUtils.setWorksheetVisibility(newWorksheet, false);		
		WorkbookUtils.protectWorkbook(workbookAutomation, true, false);
		
		return newWorksheet;
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
