/**
 * 
 */
package de.tudresden.annotator.main;

import org.eclipse.swt.SWT;
import org.eclipse.swt.events.SelectionAdapter;
import org.eclipse.swt.events.SelectionEvent;
import org.eclipse.swt.events.SelectionListener;
import org.eclipse.swt.ole.win32.OleAutomation;
import org.eclipse.swt.ole.win32.OleEvent;
import org.eclipse.swt.ole.win32.OleListener;
import org.eclipse.swt.ole.win32.Variant;
import org.eclipse.swt.widgets.Event;
import org.eclipse.swt.widgets.Listener;
import org.eclipse.swt.widgets.MessageBox;

import de.tudresden.annotator.annotations.AnnotationClass;
import de.tudresden.annotator.annotations.WorkbookAnnotation;
import de.tudresden.annotator.annotations.WorksheetAnnotation;
import de.tudresden.annotator.annotations.utils.RangeAnnotationsSheet;
import de.tudresden.annotator.annotations.utils.AnnotationHandler;
import de.tudresden.annotator.annotations.utils.AnnotationResult;
import de.tudresden.annotator.annotations.utils.AnnotationStatusSheet;
import de.tudresden.annotator.annotations.utils.ValidationResult;
import de.tudresden.annotator.oleutils.ApplicationUtils;
import de.tudresden.annotator.oleutils.RangeUtils;
import de.tudresden.annotator.oleutils.WorkbookUtils;
import de.tudresden.annotator.oleutils.WorksheetUtils;

/**
 * @author Elvis Koci
 */
public class GUIListeners {
	
	/**
	 * 
	 * @return
	 */
	protected static Listener createCloseApplicationEventListener(){
				
		return new Listener()
	    {
	        public void handleEvent(Event event)
	        {	
	        	
	    		MainWindow wm = MainWindow.getInstance();
	    		String directoryPath = wm.getDirectoryPath();
	    		String fileName = wm.getFileName();
	    		OleAutomation embeddedWorkbook = wm.getEmbeddedWorkbook();

	        	if(!wm.isControlSiteNull() && wm.isControlSiteDirty() && embeddedWorkbook!=null){
	        		int style = SWT.YES | SWT.NO | SWT.CANCEL | SWT.ICON_WARNING ;
	        		MessageBox messageBox = MainWindow.getInstance().createMessageBox(style);
	 	            messageBox.setMessage("Want to save your changes?");
	 	            
	 	            int response = messageBox.open();	 	 	            
	 	            if( response== SWT.YES){	
	 	            	
	 	            	String filePath = directoryPath+"\\"+fileName;
	 	            	boolean isSaved = FileUtils.saveProgress(embeddedWorkbook, filePath);
	 	            	
	 	            	if(!isSaved){
	 	            		int styleError = SWT.ICON_ERROR;
	 		        		MessageBox errorMessageBox = MainWindow.getInstance().createMessageBox(styleError);
	 		        		errorMessageBox.setMessage("ERROR: Could not save the file!");
	 	            		event.doit = false;
	 	            	}
	 	            	
	 	            	WorkbookUtils.closeEmbeddedWorkbook(embeddedWorkbook, false);
	 	            	MainWindow.getInstance().disposeControlSite();
	 	            	MainWindow.getInstance().disposeShell();
	 	            	event.doit = true;
	 	            } 
	 	            
	 	            if(response == SWT.NO){
	 	            	MainWindow.getInstance().disposeControlSite();
	 	            	MainWindow.getInstance().disposeShell();
	 	            	event.doit = true;
	 	            } 
	 	            
	 	            event.doit = false;
	        	}else{
	        		int style = SWT.YES | SWT.NO | SWT.ICON_QUESTION;
	        		MessageBox messageBox = MainWindow.getInstance().createMessageBox(style);
	 	            messageBox.setMessage("Do you want to close the application?");
	 	            
	 	            int response = messageBox.open();
	 	            if( response== SWT.YES){
	 	        	    MainWindow.getInstance().disposeControlSite();
	 	            	MainWindow.getInstance().disposeShell();
	 	            	event.doit = true;
	 	            }
	 	            event.doit = false;
	        	}
	        }
	    };
	}	
	
	
	/**
	 * Create a SheetSelection OLE event listener
	 * @param application
	 * @return an OleListener 
	 */
	protected static OleListener createSheetSelectionEventListener(){
		
		OleListener listener = new OleListener() {
	        public void handleEvent (OleEvent e) {
	        	
	        	Variant[] args = e.arguments;
	        	
	            /*
	             * the first argument is a Range object. Get the number and range of selected areas 
	             */	        	
	        	OleAutomation rangeAutomation = args[0].getAutomation();
	        	MainWindow.getInstance().setCurrentSelection(RangeUtils.getRangeAddress(rangeAutomation).split(","));
	        	args[0].dispose();
	        	rangeAutomation.dispose();	
	        	
				/*
				 * the second argument is a Worksheet object. Get the name and index of the worksheet.
				 */
	        	OleAutomation worksheetAutomation = args[1].getAutomation();		        
	        	MainWindow.getInstance().setActiveWorksheetName(WorksheetUtils.getWorksheetName(worksheetAutomation));
	        	MainWindow.getInstance().setActiveWorksheetIndex(WorksheetUtils.getWorksheetIndex(worksheetAutomation));
				args[1].dispose();	
				worksheetAutomation.dispose();
						
				MainWindow.getInstance().setFocusToShell();
				MainWindow.getInstance().setFocusToShell();

	        }
	    };	       
	    return listener;
	}
	
	
	/**
	 * Create a SheetActivate OLE event listener
	 * @param application
	 * @return
	 */
	protected static OleListener createSheetActivationEventListener(){
		
		OleListener listener = new OleListener() {
	        public void handleEvent (OleEvent e) {
	        	
	        	Variant[] args = e.arguments;
	        	
	        	/*
	             * This event returns only one argument, a Worksheet object. Get the name and index of the activated worksheet.
	             */ 					
				OleAutomation worksheetAutomation = args[0].getAutomation();        
				String sheetName = WorksheetUtils.getWorksheetName(worksheetAutomation);
				int sheetIndex = WorksheetUtils.getWorksheetIndex(worksheetAutomation);

				
				MainWindow.getInstance().setActiveWorksheetName(sheetName);
				MainWindow.getInstance().setActiveWorksheetIndex(sheetIndex);
				args[0].dispose();
				worksheetAutomation.dispose();
				
				// adjust the bar menu according to the properties of this worksheet
				MenuUtils.adjustBarMenuForSheet();		
								
				// return the focus to the embedded excel workbook, if it does not have it already
				MainWindow.getInstance().setFocusToControlSite();	
	        }
	    };	       
	    return listener;
	}
	
	/**
	 * 
	 * @return
	 */
	protected static Listener createArrowButtonPressedEventListener(){
		return new Listener() {
			 @Override
	         public void handleEvent(Event e) {
	        	if(e.keyCode == SWT.ARROW_UP || e.keyCode == SWT.ARROW_DOWN || 
	        	   e.keyCode == SWT.ARROW_LEFT || e.keyCode == SWT.ARROW_RIGHT)
	            {
	        		MainWindow.getInstance().setFocusToControlSite();
	            }
	        }
		};
	}
	
	/**
	 * 
	 * @return
	 */
	protected static Listener createMouseWheelEventListener(){		
		return new Listener() {
			@Override
			public void handleEvent(Event e) {
					MainWindow.getInstance().setFocusToControlSite();
			}		
		};
	}
		
	/**
	 * 
	 * @return
	 */
	protected static SelectionListener createFileOpenSelectionListener(){
		
		return new SelectionAdapter() {
			@Override
			public void widgetSelected(SelectionEvent e){
				
				// open the files selection window
				FileUtils.fileOpen();
				
				// get the OleAutomation for the embedded workbook
				OleAutomation workbookAutomation = MainWindow.getInstance().getEmbeddedWorkbook();
				if (workbookAutomation == null) {
					return; // there is no embedded workbook (file)
				}
				
				
				// clear all existing annotations in memory structure, 
				// if they exist from the previously opened file 
				AnnotationHandler.getWorkbookAnnotation().removeAllAnnotations();
		
				// create the base in memory structure for storing annotation data
				// AnnotationHandler.createBaseAnnotations(workbookAutomation);
				
				AnnotationStatusSheet.readAnnotationStatuses(workbookAutomation);
				
				// read the annotation data and recreate in memory structure
				RangeAnnotationsSheet.readRangeAnnotations(workbookAutomation);
				
				// re-draw all the annotation in memory structure 
				AnnotationHandler.drawAllAnnotations(workbookAutomation);								
				
				// adjust the menu items in the menu bar for the file that was just openned
				MenuUtils.adjustBarMenuForOpennedFile();
			}
		};
	}
		
	/**
	 * 
	 * @return
	 */
	protected static SelectionListener createFileSaveSelectionListener(){
		
		return new SelectionAdapter() {
			@Override
			public void widgetSelected(SelectionEvent e) {
				
				OleAutomation embeddedWorkbook = MainWindow.getInstance().getEmbeddedWorkbook();
				String fileName = MainWindow.getInstance().getFileName();
				String directory = MainWindow.getInstance().getDirectoryPath();
				String filePath = directory+"\\"+fileName;
				
				OleAutomation applicationAutomation = WorkbookUtils.getApplicationAutomation(embeddedWorkbook);
				
				ApplicationUtils.setDisplayAlerts(applicationAutomation, "False");		
				boolean result = FileUtils.saveProgress(embeddedWorkbook, filePath);
				if(result){		
					// TODO: Mark Annotated Files. 
            		// FileUtils.markFileAsAnnotated(directory, fileName, 1);
            		int style = SWT.ICON_INFORMATION;
					MessageBox message = MainWindow.getInstance().createMessageBox(style);
					message.setMessage("The file was successfully saved.");
					message.open();
				}else{
					int style = SWT.ICON_ERROR;
					MessageBox message = MainWindow.getInstance().createMessageBox(style);
					message.setMessage("ERROR: The file could not be saved!");
					message.open();
				}				
				ApplicationUtils.setDisplayAlerts(applicationAutomation, "True");	
			}
		};
	}
	
	/**
	 * 
	 * @return
	 */
	protected static SelectionListener createFileCloseSelectionListener(){
		
		return new SelectionAdapter() {
			@Override
			public void widgetSelected(SelectionEvent e) {
				
				if( !MainWindow.getInstance().isControlSiteNull() && 
						MainWindow.getInstance().isControlSiteDirty() &&
						 	MainWindow.getInstance().getEmbeddedWorkbook()!=null){
									
	        		int style = SWT.YES | SWT.NO | SWT.CANCEL | SWT.ICON_WARNING ;
	        		MessageBox messageBox = MainWindow.getInstance().createMessageBox(style);
	 	            messageBox.setMessage("Want to save your changes?");
	 	            
	 	            int response = messageBox.open();	 	 	            
	 	            if( response== SWT.YES){	
	 	            	OleAutomation embeddedWorkbook = MainWindow.getInstance().getEmbeddedWorkbook();
	 	            	String filePath =  MainWindow.getInstance().getDirectoryPath()+"\\"+MainWindow.getInstance().getFileName();
	 	            	boolean isSaved = FileUtils.saveProgress(embeddedWorkbook, filePath);
	 	            	
	 	            	if(!isSaved){
	 	            		int messageStyle = SWT.ICON_ERROR;
	 						MessageBox message = MainWindow.getInstance().createMessageBox(messageStyle);
	 						message.setMessage("ERROR: The file could not be saved!");
	 						message.open();
	 	            	}else{					
	 	            		//String directory = MainWindow.getInstance().getDirectoryPath();
	 	            		//String fileName = MainWindow.getInstance().getFileName();
	 	            		//FileUtils.markFileAsAnnotated(directory, fileName, 1);
	 	            	}
	 	            } 
	 	            
	 	            if(response == SWT.NO || response == SWT.YES){
	 	            	OleAutomation embeddedWorkbook  = MainWindow.getInstance().getEmbeddedWorkbook();
	 					WorkbookUtils.closeEmbeddedWorkbook(embeddedWorkbook, false);
	 					MainWindow.getInstance().disposeControlSite();
	 	            } 
				}else{
					
	        		int style = SWT.YES | SWT.NO | SWT.ICON_QUESTION;
	        		MessageBox messageBox = MainWindow.getInstance().createMessageBox(style);
	 	            messageBox.setMessage("Are you sure you want to close the file?");
	 	            
	 	            int response = messageBox.open();
	 	            if( response== SWT.YES){
	 	            	OleAutomation embeddedWorkbook  = MainWindow.getInstance().getEmbeddedWorkbook();
	 					WorkbookUtils.closeEmbeddedWorkbook(embeddedWorkbook, false);
	 					MainWindow.getInstance().disposeControlSite();
	 	            }
	        	}			
			}
		};
	}
	
	/**
	 * 
	 * @return
	 */
	protected static SelectionListener createExportAsCSVSelectionListener(){
		
		return new SelectionAdapter() {
			@Override
			public void widgetSelected(SelectionEvent e) {
				OleAutomation workbookAutomation = MainWindow.getInstance().getEmbeddedWorkbook();
				String directoryPath = MainWindow.getInstance().getDirectoryPath();
				String fileName = MainWindow.getInstance().getFileName();				
				boolean isSuccess = RangeAnnotationsSheet.exportRangeAnnotationsAsCSV(workbookAutomation, directoryPath, fileName);
				
				if(isSuccess){
					MessageBox messageBox = MainWindow.getInstance().createMessageBox(SWT.ICON_INFORMATION);
					messageBox.setText("Information");
		            messageBox.setMessage("The annotation data were successfully exported at:\n"+directoryPath);
		            messageBox.open();
				}else{
					MessageBox messageBox = MainWindow.getInstance().createMessageBox(SWT.ICON_ERROR);
					messageBox.setText("Error Message");
		            messageBox.setMessage("An error occured. Could not export the annotation data as csv.");
		            messageBox.open();
				}
			}
		};
	}
	
	
	/**
	 * 
	 * @return
	 */
	protected static SelectionListener createExportAsWorkbookSelectionListener(){
		
		return new SelectionAdapter() {
			@Override
			public void widgetSelected(SelectionEvent e) {
					MessageBox messageBox = MainWindow.getInstance().createMessageBox(SWT.ICON_INFORMATION);
					messageBox.setText("Information");
		            messageBox.setMessage("This option is not implemented yet");
			}
		};
	}
	
	/**
	 * 
	 * @return
	 */
	protected static SelectionListener createFileExitSelectionListener(){
		
		return new SelectionAdapter() {
			@Override
			public void widgetSelected(SelectionEvent e) {	
				
				if( !MainWindow.getInstance().isControlSiteNull() && 
						MainWindow.getInstance().isControlSiteDirty() &&
						 	MainWindow.getInstance().getEmbeddedWorkbook()!=null){
								
	        		int style = SWT.YES | SWT.NO | SWT.CANCEL | SWT.ICON_WARNING ;
	        		MessageBox messageBox = MainWindow.getInstance().createMessageBox(style);
	 	            messageBox.setMessage("Want to save your changes?");
	 	            
	 	            int response = messageBox.open();	 	 	            
	 	            if( response== SWT.YES){	
	 	            	OleAutomation embeddedWorkbook = MainWindow.getInstance().getEmbeddedWorkbook();
	 	            	String filePath =  MainWindow.getInstance().getDirectoryPath()+"\\"+MainWindow.getInstance().getFileName();
	 	            	boolean isSaved = FileUtils.saveProgress(embeddedWorkbook, filePath);
	 	            	
	 	            	if(!isSaved){
	 	            		int messageStyle = SWT.ICON_ERROR;
	 						MessageBox message = MainWindow.getInstance().createMessageBox(messageStyle);
	 						message.setMessage("ERROR: The file could not be saved!");
	 						message.open();
	 	            	}else{
	 	            		//String directory = MainWindow.getInstance().getDirectoryPath();
	 	            		//String fileName = MainWindow.getInstance().getFileName();
	 	            		//FileUtils.markFileAsAnnotated(directory, fileName, 1);
	 	            		
		 	            	WorkbookUtils.closeEmbeddedWorkbook(embeddedWorkbook, false);
		 	            	MainWindow.getInstance().disposeControlSite();
		 	            	MainWindow.getInstance().disposeShell();
	 	            	}
	 	            } 
	 	            
	 	            if(response == SWT.NO){
	 	            	MainWindow.getInstance().disposeControlSite();
	 	            	MainWindow.getInstance().disposeShell();
	 	            } 
	        	}else{
	        		int style = SWT.YES | SWT.NO | SWT.ICON_QUESTION;
	        		MessageBox messageBox = MainWindow.getInstance().createMessageBox(style);
	 	            messageBox.setMessage("Do you want to close the application?");
	 	            
	 	            int response = messageBox.open();
	 	            if( response== SWT.YES){
	 	        	    MainWindow.getInstance().disposeControlSite();
	 	            	MainWindow.getInstance().disposeShell();
	 	            }
	        	}
			}
		};
	}
	
	/**
	 * 
	 * @return
	 */
	protected static SelectionListener createSheetCompletedSelectionListener(){
		return new SelectionAdapter() {
			@Override
			public void widgetSelected(SelectionEvent e){
				String sheetName = MainWindow.getInstance().getActiveWorksheetName();
				WorkbookAnnotation workbookAnnotation = AnnotationHandler.getWorkbookAnnotation();
				WorksheetAnnotation  sheetAnnotation = workbookAnnotation.getWorksheetAnnotations().get(sheetName);
				
				sheetAnnotation.setCompleted(!sheetAnnotation.isCompleted());			
				MenuUtils.adjustBarMenuForSheet();
				
				int style = SWT.ICON_INFORMATION;
				MessageBox mb = MainWindow.getInstance().createMessageBox(style);
				String value = String.valueOf((sheetAnnotation.isCompleted())).toUpperCase();
				mb.setMessage(" Sheet status was updated to Completed := "+value); 
				mb.open();
			}
		};
	}
	
	/**
	 * 
	 * @return
	 */
	protected static SelectionListener createSheetNotApplicableSelectionListener(){
		return new SelectionAdapter() {
			@Override
			public void widgetSelected(SelectionEvent e){
				String sheetName = MainWindow.getInstance().getActiveWorksheetName();
				WorkbookAnnotation workbookAnnotation = AnnotationHandler.getWorkbookAnnotation();
				WorksheetAnnotation  sheetAnnotation = workbookAnnotation.getWorksheetAnnotations().get(sheetName);
				
				sheetAnnotation.setNotApplicable(!sheetAnnotation.isNotApplicable());
				MenuUtils.adjustBarMenuForSheet();
				
				int style = SWT.ICON_INFORMATION;
				MessageBox mb = MainWindow.getInstance().createMessageBox(style);
				String value = String.valueOf((sheetAnnotation.isNotApplicable())).toUpperCase();
				mb.setMessage(" Sheet status was updated to NotApplicable := "+value); 
				mb.open();
			}
		};
	}
	
	/**
	 * 
	 * @return
	 */
	protected static SelectionListener createFileCompletedSelectionListener(){
		return new SelectionAdapter() {
			@Override
			public void widgetSelected(SelectionEvent e){
				
				WorkbookAnnotation workbookAnnotation = AnnotationHandler.getWorkbookAnnotation();		
				workbookAnnotation.setCompleted(!workbookAnnotation.isCompleted());
				MenuUtils.adjustBarMenuForWorkbook();
				
				int style = SWT.ICON_INFORMATION;
				MessageBox mb = MainWindow.getInstance().createMessageBox(style);
				String value = String.valueOf((workbookAnnotation.isCompleted())).toUpperCase();
				mb.setMessage("File (Workbook) status was updated to Completed := "+value); 
				mb.open();
			}
		};
	}
	
	/**
	 * 
	 * @return
	 */
	protected static SelectionListener createFileNotApplicableSelectionListener(){
		return new SelectionAdapter() {
			@Override
			public void widgetSelected(SelectionEvent e){
				
				WorkbookAnnotation workbookAnnotation = AnnotationHandler.getWorkbookAnnotation();		
				workbookAnnotation.setNotApplicable(!workbookAnnotation.isNotApplicable());
				MenuUtils.adjustBarMenuForWorkbook();
				
				int style = SWT.ICON_INFORMATION;
				MessageBox mb = MainWindow.getInstance().createMessageBox(style);
				String value = String.valueOf((workbookAnnotation.isNotApplicable())).toUpperCase();
				mb.setMessage("File (Workbook) status was updated to NotApplicable := "+value); 
				mb.open();
			}
		};
	}
	
	
	/**
	 * 
	 * @param annotationClass
	 * @return
	 */
	protected static SelectionListener createRangeAnnotationSelectionListener(AnnotationClass annotationClass){
		
		return new SelectionAdapter() {
			@Override
			public void widgetSelected(SelectionEvent e) {
				 
				 OleAutomation workbookAutomation = MainWindow.getInstance().getEmbeddedWorkbook();
				 
				 String sheetName = MainWindow.getInstance().getActiveWorksheetName();
				 int sheetIndex = MainWindow.getInstance().getActiveWorksheetIndex();
				 String[] currentSelection = MainWindow.getInstance().getCurrentSelection();

				 AnnotationResult  result=  
						 AnnotationHandler.annotate(workbookAutomation, sheetName, sheetIndex,   
				 		 currentSelection, annotationClass);				 
				 
				 if(result.getValidationResult()!=ValidationResult.OK){
	        		MessageBox messageBox = MainWindow.getInstance().createMessageBox(SWT.ICON_ERROR);
	 	            messageBox.setMessage(result.getMessage());
	 	            messageBox.open();
				 }
				 				 
				 MainWindow.getInstance().setFocusToShell();
				 // MainWindow.getInstance().setFocusToControlSite();				 
			}
		};
	}
	
	/**
	 * 
	 * @return
	 */
	protected static SelectionListener createHideAllAnnotationsSelectionListener(){
		return new SelectionAdapter() {
			@Override
			public void widgetSelected(SelectionEvent e) {
				OleAutomation workbookAutomation = MainWindow.getInstance().getEmbeddedWorkbook();	
				AnnotationHandler.setVisilityForShapeAnnotations(workbookAutomation, false);
			}
		};
	}
	
	/**
	 * 
	 * @return
	 */
	protected static SelectionListener createHideInSheetAnnotationsSelectionListener(){
		return new SelectionAdapter() {
			@Override
			public void widgetSelected(SelectionEvent e) {
				OleAutomation workbookAutomation = MainWindow.getInstance().getEmbeddedWorkbook();
				String sheetName = MainWindow.getInstance().getActiveWorksheetName();
				AnnotationHandler.setVisibilityForWorksheetShapeAnnotations(workbookAutomation, sheetName, false);
			}
		};
	}
	
	/**
	 * 
	 * @return
	 */
	protected static SelectionListener createDeleteAllAnnotationsSelectionListener(){
		return new SelectionAdapter() {
			@Override
			public void widgetSelected(SelectionEvent e) {
				OleAutomation workbookAutomation = MainWindow.getInstance().getEmbeddedWorkbook();	

				AnnotationHandler.deleteAllShapeAnnotations(workbookAutomation);
				
				WorkbookAnnotation workbookAnnotation = AnnotationHandler.getWorkbookAnnotation();
				workbookAnnotation.removeAllAnnotations();
				
				RangeAnnotationsSheet.deleteAllRangeAnnotations(workbookAutomation);
			}
		};
	}
	
	/**
	 * 
	 * @return
	 */
	protected static SelectionListener createDeleteInSheetAnnotationsSelectionListener(){
		return new SelectionAdapter() {
			@Override
			public void widgetSelected(SelectionEvent e) {

				OleAutomation workbookAutomation = MainWindow.getInstance().getEmbeddedWorkbook();
				String sheetName = MainWindow.getInstance().getActiveWorksheetName();
				
				AnnotationHandler.deleteShapeAnnotationsFromWorksheet(workbookAutomation, sheetName);
				
				WorkbookAnnotation workbookAnnotation = AnnotationHandler.getWorkbookAnnotation();
				workbookAnnotation.removeAllAnnotationsFromSheet(sheetName);
				
				RangeAnnotationsSheet.deleteRangeAnnotationsForWorksheet(workbookAutomation, sheetName, false);		
			}
		};
	}
}
