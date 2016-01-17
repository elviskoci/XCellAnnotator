/**
 * 
 */
package de.tudresden.annotator.annotations.utils;

import java.util.Collection;

import org.eclipse.swt.SWT;
import org.eclipse.swt.ole.win32.OleAutomation;
import org.eclipse.swt.widgets.MessageBox;

import de.tudresden.annotator.annotations.WorkbookAnnotation;
import de.tudresden.annotator.annotations.WorksheetAnnotation;
import de.tudresden.annotator.main.MainWindow;
import de.tudresden.annotator.oleutils.CollectionsUtils;
import de.tudresden.annotator.oleutils.RangeUtils;
import de.tudresden.annotator.oleutils.WorkbookUtils;
import de.tudresden.annotator.oleutils.WorksheetUtils;

/**
 * @author Elvis Koci
 */
public class AnnotationStatusSheet {
	
	protected static final String name = "Annotation_Status_Data";
	private static final String startColumnChar = "A";
	private static final int startColumnIndex = 1;
	private static final int startRow = 1; 
	
	/**
	 * Save the annotation status for the workbook and worksheets
	 * @param workbookAutomation an OleAutomation to access the embedded workbook
	 */
	public static void saveAnnotationStatuses(OleAutomation workbookAutomation){
			
		OleAutomation annotationStatusSheet =  WorkbookUtils.getWorksheetAutomationByName(workbookAutomation, name);
		
		if(annotationStatusSheet==null){		
			annotationStatusSheet = createAnnotationStatusSheet(workbookAutomation);	
		}else{
			
			WorksheetUtils.unprotectWorksheet(annotationStatusSheet);
			
			// delete all the existing data from the worksheet. 
			// by removing all existing data we ensure that the "new" data will have
			// the right format. So, they are not effected by the existing data.
			OleAutomation usedRange = WorksheetUtils.getUsedRange(annotationStatusSheet);		
			RangeUtils.deleteRange(usedRange);	
			usedRange.dispose();
			
			// re-create the header
			createHeaderRow(annotationStatusSheet);
			WorksheetUtils.protectWorksheet(annotationStatusSheet);	
		}
		
		writeStatuses(annotationStatusSheet);		
		annotationStatusSheet.dispose();
	}
	
	
	public static void readAnnotationStatuses(OleAutomation workbookAutomation){
		
		OleAutomation annotationStatusSheet =  WorkbookUtils.getWorksheetAutomationByName(workbookAutomation, name);
		AnnotationHandler.createBaseAnnotations(workbookAutomation);
		if(annotationStatusSheet==null){		
			return;
		}
		
		if(!validateAnnotationStatusSheet(annotationStatusSheet)){
			return;
		}
		
		if(!readWorkbookAnnotationStatus(annotationStatusSheet)){
			return;
		}
		
		readWorksheetAnnotationsStatuses(annotationStatusSheet);
	}
	
	
	
	/**
	 * Create the sheet that will store the status of the annotation for the workbook and the individual worksheets 
	 * @param workbookAutomation an OleAutomation to access the embedded workbook
	 * @return the OleAutomation of the created Annotation_Status sheet
	 */
	private static OleAutomation createAnnotationStatusSheet(OleAutomation workbookAutomation){
		
		WorkbookUtils.unprotectWorkbook(workbookAutomation);
		
		OleAutomation annotationStatusSheet = WorkbookUtils.addWorksheetAsLast(workbookAutomation);
		WorksheetUtils.setWorksheetName(annotationStatusSheet, name);
		
		createHeaderRow(annotationStatusSheet);
		
		WorksheetUtils.setWorksheetVisibility(annotationStatusSheet, false);
		WorkbookUtils.protectWorkbook(workbookAutomation, true, false);
		
		return annotationStatusSheet;
	}
	
	
	/**
	 * Create (write) the header row that contains the field names 
	 * @param annotationStatusSheet  an OleAutomation that provides access to the sheet that maintains annotation status data
	 */
	private static void createHeaderRow(OleAutomation annotationStatusSheet){
		
		OleAutomation field1 = WorksheetUtils.getCell(annotationStatusSheet, startRow, startColumnIndex);
		RangeUtils.setValue(field1, "Name");
		field1.dispose();
		
		OleAutomation field2 = WorksheetUtils.getCell(annotationStatusSheet, startRow, startColumnIndex+1);
		RangeUtils.setValue(field2, "Completed");
		field2.dispose();
		
		OleAutomation field3 = WorksheetUtils.getCell(annotationStatusSheet, startRow, startColumnIndex+2);
		RangeUtils.setValue(field3, "NotApplicable");
		field3.dispose();
	}
	
	
	/**
	 * Write new row of annotation status data
	 * @param annotationStatusSheet an OleAutomation that provides access to the sheet that maintains annotation status data
	 * @param row an integer that represents the index of the row to write the data
	 * @param name a string that represents the name of the worksheet or workbook
	 * @param isCompleted a boolean value that specifies if the annotation (worksheet or workbook) is completed or not
	 * @param isNotApplicable a boolean value that specifies if the annotation (worksheet or workbook) is applicable or not
	 */
	private static void writeNewDataRow(OleAutomation annotationStatusSheet, int row, String name, 
														boolean isCompleted, boolean isNotApplicable){		
		
		WorksheetUtils.unprotectWorksheet(annotationStatusSheet);
		
		OleAutomation field1 = WorksheetUtils.getCell(annotationStatusSheet, startRow, startColumnIndex);
		RangeUtils.setValue(field1, name);
		field1.dispose();
		
		OleAutomation field2 = WorksheetUtils.getCell(annotationStatusSheet, startRow, startColumnIndex+1);
		RangeUtils.setValue(field2, String.valueOf(isCompleted));
		field2.dispose();
		
		OleAutomation field3 = WorksheetUtils.getCell(annotationStatusSheet, startRow, startColumnIndex+2);
		RangeUtils.setValue(field3, String.valueOf(isNotApplicable));
		field3.dispose();
		
		WorksheetUtils.protectWorksheet(annotationStatusSheet);	
	}
	
	/**
	 * This method is used to write the status data of many worksheet annotations at once.
	 * @param annotationStatusSheet an OleAutomation that provides access to the sheet that maintains the annotation status data
	 * @param row an integer that represents the index of the row from where to start writing the data
	 * @param worksheetAnnotations an array of WorksheetAnnotations
	 */
	private static void writeManyWorksheetAnnotationStatuses(OleAutomation annotationStatusSheet, int row, WorksheetAnnotation[] worksheetAnnotations){
		
		WorksheetUtils.unprotectWorksheet(annotationStatusSheet);
		
		int rowIndex = row;
		for (WorksheetAnnotation worksheetAnnotation : worksheetAnnotations) {
			
			OleAutomation field1 = WorksheetUtils.getCell(annotationStatusSheet, rowIndex, startColumnIndex);
			RangeUtils.setValue(field1, worksheetAnnotation.getSheetName());
			
			OleAutomation field2 = WorksheetUtils.getCell(annotationStatusSheet, rowIndex, startColumnIndex+1);
			RangeUtils.setValue(field2, String.valueOf(worksheetAnnotation.isCompleted()));
			
			OleAutomation field3 = WorksheetUtils.getCell(annotationStatusSheet, rowIndex, startColumnIndex+2);
			RangeUtils.setValue(field3, String.valueOf(worksheetAnnotation.isNotApplicable()));
			
			rowIndex++;
		}
		
		WorksheetUtils.protectWorksheet(annotationStatusSheet);
	}
	
	/**
	 * Write the status data of the Workbook and Worksheet annotations for the embedded excel file 
	 * @param annotationStatusSheet an OleAutomation that provides access to the sheet that maintains the annotation status data
	 */
	private static void writeStatuses(OleAutomation annotationStatusSheet){
		
		WorksheetUtils.unprotectWorksheet(annotationStatusSheet);
		
		// write the annotation status (data) of the workbook annotation
		WorkbookAnnotation workbookAnnotation = AnnotationHandler.getWorkbookAnnotation();

		OleAutomation bookName = WorksheetUtils.getCell(annotationStatusSheet, startRow+1, startColumnIndex);
		RangeUtils.setValue(bookName, "Workbook");
		bookName.dispose();
		
		OleAutomation isBookCompleted = WorksheetUtils.getCell(annotationStatusSheet, startRow+1, startColumnIndex+1);
		RangeUtils.setValue(isBookCompleted, String.valueOf(workbookAnnotation.isCompleted()));
		isBookCompleted.dispose();
		
		OleAutomation isBookNotApplicable = WorksheetUtils.getCell(annotationStatusSheet, startRow+1, startColumnIndex+2);
		RangeUtils.setValue(isBookNotApplicable, String.valueOf(workbookAnnotation.isNotApplicable()));
		isBookNotApplicable.dispose();
				
		// write the annotation status (data) of each worksheet annotation
		Collection<WorksheetAnnotation> collection = workbookAnnotation.getWorksheetAnnotations().values(); 
		WorksheetAnnotation[] worksheetAnnotations = collection.toArray(new WorksheetAnnotation[collection.size()]);
		
		int rowIndex = startRow+2;
		for (WorksheetAnnotation worksheetAnnotation : worksheetAnnotations) {
			
			OleAutomation sheetName = WorksheetUtils.getCell(annotationStatusSheet, rowIndex, startColumnIndex);
			RangeUtils.setValue(sheetName, worksheetAnnotation.getSheetName());
			
			OleAutomation isSheetCompleted = WorksheetUtils.getCell(annotationStatusSheet, rowIndex, startColumnIndex+1);
			RangeUtils.setValue(isSheetCompleted, String.valueOf(worksheetAnnotation.isCompleted()));
			
			OleAutomation isSheetNotApplicable = WorksheetUtils.getCell(annotationStatusSheet, rowIndex, startColumnIndex+2);
			RangeUtils.setValue(isSheetNotApplicable, String.valueOf(worksheetAnnotation.isNotApplicable()));
			
			rowIndex++;
		}
		
		WorksheetUtils.protectWorksheet(annotationStatusSheet);
	}
	
	/**
	 * Validate the "Annotation_Status" sheet to ensure that the data are in the expected format. 
	 * @param annotationStatusSheet an OleAutomation that provides access to the sheet that maintains the annotation status data
	 * @return true if the status data are in the expected format, false otherwise
	 */
	private static boolean validateAnnotationStatusSheet(OleAutomation annotationStatusSheet){
		
		OleAutomation usedRange = WorksheetUtils.getUsedRange(annotationStatusSheet);
		if(usedRange==null){
			int style = SWT.ICON_WARNING;
			MessageBox message = MainWindow.getInstance().createMessageBox(style);
			message.setMessage("Could not recover the annotation status from the previous session. "
					+ "The \"Annotation_Status\" sheet is empty.");
			message.open();
			return false;
		}
		
		
		OleAutomation rangeColumns = RangeUtils.getRangeColumns(usedRange);
		int countColumns = CollectionsUtils.countItemsInCollection(rangeColumns);
		rangeColumns.dispose();
		
		if(countColumns!=3){
			usedRange.dispose();
						
			int style = SWT.ICON_WARNING;
			MessageBox message = MainWindow.getInstance().createMessageBox(style);
			message.setMessage("Could not recover the annotation status from the previous session. "
					+ "The annotation status data are not in the expected format. "
					+ "The expected number of columns is 3. Instead, the application found "+countColumns+" column/s.");
			message.open();	
			return false;
		}
				
		OleAutomation rangeRows = RangeUtils.getRangeRows(usedRange);
		OleAutomation headerRow = CollectionsUtils.getItemByIndex(rangeRows, 1, false);
		rangeRows.dispose();
		usedRange.dispose();
		
		String values[] = RangeUtils.getRangeValues(headerRow);
		
		if(!(values[0].compareToIgnoreCase("Name")==0 &&
			values[1].compareToIgnoreCase("Completed")==0 &&
			values[2].compareToIgnoreCase("NotApplicable")==0)){
			
			int style = SWT.ICON_WARNING;
			MessageBox message = MainWindow.getInstance().createMessageBox(style);
			message.setMessage("Could not recover the annotation status from the previous session. "
					+ "The header row does not contain the expected fields.");
			message.open();
			return false;
		}
		
		return true;
	}
	
	
	/**
	 * Read the workbook annotation status data
	 * @param annotationStatusSheet an OleAutomation that provides access to the sheet that maintains the annotation status data
	 * @return true if the status data were successfully read, false otherwise
	 */
	private static boolean readWorkbookAnnotationStatus(OleAutomation annotationStatusSheet){
		
		WorkbookAnnotation wa = AnnotationHandler.getWorkbookAnnotation();
		String topLeftCell = "$"+startColumnChar+"$"+(startRow+1);
		OleAutomation topRightCellAuto = WorksheetUtils.getCell(annotationStatusSheet, (startRow+1), startColumnIndex+2);
		String downLeftCell = RangeUtils.getRangeAddress(topRightCellAuto);
		topRightCellAuto.dispose();
		
		OleAutomation workbookStatusRow = WorksheetUtils.getRangeAutomation(annotationStatusSheet, topLeftCell, downLeftCell);
		String[] values= RangeUtils.getRangeValues(workbookStatusRow);
		workbookStatusRow.dispose();
		
		if(!validateRowData(values)){
			return false;
		}
		
		boolean isCompleted = Integer.valueOf(values[1])==1;
		wa.setCompleted(isCompleted);
		boolean isNotApplicable = Integer.valueOf(values[2])==1;
		wa.setNotApplicable(isNotApplicable);
			
		return true;	
	}

	/**
	 * Read the status data for the worksheet annotations
	 * @param annotationStatusSheet an OleAutomation that provides access to the sheet that maintains the annotation status data
	 * @return true if the status data were successfully read, false otherwise
	 */
	private static boolean readWorksheetAnnotationsStatuses(OleAutomation annotationStatusSheet){
		
		OleAutomation usedRange = WorksheetUtils.getUsedRange(annotationStatusSheet);		
		OleAutomation rangeRows = RangeUtils.getRangeRows(usedRange);
		int countRows = CollectionsUtils.countItemsInCollection(rangeRows);		
		usedRange.dispose();
		
		OleAutomation topRightCellAuto = WorksheetUtils.getCell(annotationStatusSheet, (startRow), startColumnIndex+2);
		String topRightCell= RangeUtils.getRangeAddress(topRightCellAuto);
		topRightCellAuto.dispose();
		String endColumnChar = topRightCell.replaceAll("[0-9\\$]+",""); 
		
		int rowIndex = startRow+2;

		WorkbookAnnotation wa = AnnotationHandler.getWorkbookAnnotation();
		while(rowIndex<=countRows) {
			
			String startCell = "$"+startColumnChar+"$"+rowIndex;
			String endCell = "$"+endColumnChar+"$"+rowIndex;
			OleAutomation sheetStatusRow = WorksheetUtils.getRangeAutomation(annotationStatusSheet, startCell, endCell);
			String[] values= RangeUtils.getRangeValues(sheetStatusRow);

			if(!validateRowData(values)){
				return false;
			}
			
			String sheetName = values[0];
			WorksheetAnnotation sheetAnnotation = wa.getWorksheetAnnotations().get(sheetName);
			if(sheetAnnotation!=null){
				boolean isCompleted = Integer.valueOf(values[1])==-1;
				sheetAnnotation.setCompleted(isCompleted);
				boolean isNotApplicable = Integer.valueOf(values[2])==-1;
				sheetAnnotation.setNotApplicable(isNotApplicable);
			}
			
			rowIndex++;		
		}		
		return true;
	}
	
	/**
	 * Validate the row data. 
	 * @param data an array of string values that were read from the worksheet range - row
	 * @return true if the data pass all the validation tests, false otherwise.
	 */
	private static boolean validateRowData(String data[]){
		
		if(data.length!=3){
			int style = SWT.ICON_WARNING;
			MessageBox message = MainWindow.getInstance().createMessageBox(style);
			message.setMessage("Could not recover the annotation status from the previous session. "
					+ "The data are not in the expected format!");
			message.open();
			return false;
		}
		
		if(!(data[1].compareToIgnoreCase("0")==0  || data[1].compareToIgnoreCase("-1")==0)){
	
			int style = SWT.ICON_WARNING;
			MessageBox message = MainWindow.getInstance().createMessageBox(style);
			message.setMessage("Could not recover the annotation status from the previous session. "
					+ "Could not recognize one or more values of the field \"isCompleted\"");
			message.open();
			return false;
		}
		
		if(!(data[2].compareToIgnoreCase("0")==0  || data[2].compareToIgnoreCase("-1")==0)){
	
			int style = SWT.ICON_WARNING;
			MessageBox message = MainWindow.getInstance().createMessageBox(style);
			message.setMessage("Could not recover the annotation status from the previous session. "
					+ "Could not recognize one or more values of the field \"isNotApplicable\"");
			message.open();
			return false;
		}
		
		return true;
	}
	
	/**
	 * @return the name
	 */
	public static String getName() {
		return name;
	}
		
}
