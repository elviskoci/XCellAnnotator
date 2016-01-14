/**
 * 
 */
package de.tudresden.annotator.main;

import java.util.HashMap;

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
import de.tudresden.annotator.annotations.utils.AnnotationDataSheet;
import de.tudresden.annotator.annotations.utils.AnnotationHandler;
import de.tudresden.annotator.annotations.utils.AnnotationResult;
import de.tudresden.annotator.annotations.utils.ValidationResult;
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
		
		MainWindow wm = MainWindow.getInstance();
		String directoryPath = MainWindow.getInstance().getDirectoryPath();
		String fileName = MainWindow.getInstance().getFileName();
		OleAutomation embeddedWorkbook = MainWindow.getInstance().getEmbeddedWorkbook();
		
		return new Listener()
	    {
	        public void handleEvent(Event event)
	        {	
	        	if(!wm.isControlSiteNull() && wm.isControlSiteDirty()){
	        		int style = SWT.YES | SWT.NO | SWT.CANCEL | SWT.ICON_WARNING ;
	        		MessageBox messageBox = MainWindow.getInstance().createMessageBox(style);
	 	            messageBox.setMessage("Want to save your changes?");
	 	            
	 	            int response = messageBox.open();	 	 	            
	 	            if( response== SWT.YES){	
	 	            	
	 	            	String filePath = directoryPath+"\\"+fileName;
	 	            	boolean isSaved = FileUtils.saveProgress(embeddedWorkbook, filePath);
	 	            	
	 	            	if(!isSaved){
	 	            		System.err.println("Could not save progress!");
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
					return; // there is no embedded workbook
				}
				
				// create the base in memory structure for storing annotation data
				AnnotationHandler.createBaseAnnotations(workbookAutomation);
						
				// check if there is a "Annotation Data" sheet
				// if yes read the annotation data stored in this sheet
				// and update the in memory structure. Also, re-draw all 
				// range annotations in their corresponding sheets. 
				OleAutomation  annotationDataSheet = 
					WorkbookUtils.getWorksheetAutomationByName(workbookAutomation, AnnotationDataSheet.getName()); 
				
				if(annotationDataSheet!=null){
					
					// read the annotation data and recreate in memory structure
					AnnotationDataSheet.readAnnotationData(workbookAutomation);
					
					// re-draw all the annotation in memory structure 
					AnnotationHandler.drawAllAnnotations(workbookAutomation);								
					
					// protect the annotation data sheet if it is not protected already
					boolean isProtected = WorksheetUtils.protectWorksheet(annotationDataSheet);
					if(!isProtected){
						int style = SWT.ERROR;
						MessageBox message = MainWindow.getInstance().createMessageBox(style);
						message.setMessage("ERROR: Could not protect annotation data sheet!");
						message.open();
						WorkbookUtils.closeEmbeddedWorkbook(workbookAutomation, false);
						MainWindow.getInstance().disposeControlSite();
						return;
					}
		
					// show the annotation data sheet
					// TODO: Check why it does not make visible
					WorksheetUtils.setWorksheetVisibility(annotationDataSheet, true);
				}
						
				MenuUtils.adjustBarMenuForOpennedFile();
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
				 				 
				 MainWindow.getInstance().setFocusToControlSite();				 
			}
		};
	}
}
