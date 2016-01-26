/**
 * 
 */
package de.tudresden.annotator.main;

import java.util.Collection;
import java.util.Iterator;

import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.eclipse.swt.SWT;
import org.eclipse.swt.events.SelectionAdapter;
import org.eclipse.swt.events.SelectionEvent;
import org.eclipse.swt.events.SelectionListener;
import org.eclipse.swt.graphics.Color;
import org.eclipse.swt.ole.win32.OleAutomation;
import org.eclipse.swt.ole.win32.OleEvent;
import org.eclipse.swt.ole.win32.OleListener;
import org.eclipse.swt.ole.win32.Variant;
import org.eclipse.swt.widgets.Display;
import org.eclipse.swt.widgets.Event;
import org.eclipse.swt.widgets.Listener;
import org.eclipse.swt.widgets.MessageBox;

import de.tudresden.annotator.annotations.AnnotationClass;
import de.tudresden.annotator.annotations.RangeAnnotation;
import de.tudresden.annotator.annotations.WorkbookAnnotation;
import de.tudresden.annotator.annotations.WorksheetAnnotation;
import de.tudresden.annotator.annotations.utils.AnnotationHandler;
import de.tudresden.annotator.annotations.utils.AnnotationStatusSheet;
import de.tudresden.annotator.annotations.utils.RangeAnnotationsSheet;
import de.tudresden.annotator.oleutils.RangeUtils;
import de.tudresden.annotator.oleutils.WorkbookUtils;
import de.tudresden.annotator.oleutils.WorksheetUtils;

/**
 * @author Elvis Koci
 */
public class GUIListeners {
	
	private static final Logger logger = LogManager.getLogger(GUIListeners.class.getName());
	
	/**
	 * 
	 * @return
	 */
	protected static Listener createCloseApplicationEventListener(){
				
		return new Listener()
	    {
	        public void handleEvent(Event event)
	        {	
	        	if(!Launcher.getInstance().isControlSiteNull() && 
						AnnotationHandler.getWorkbookAnnotation().hashCode()!= AnnotationHandler.getOldWorkbookAnnotationHash()){
	    			
		    		Launcher wm = Launcher.getInstance();
		    		String directoryPath = wm.getDirectoryPath();
		    		String fileName = wm.getFileName();
		    		OleAutomation embeddedWorkbook = wm.getEmbeddedWorkbook();

	        		int style = SWT.YES | SWT.NO | SWT.CANCEL | SWT.ICON_WARNING ;
	        		MessageBox messageBox = Launcher.getInstance().createMessageBox(style);
	 	            messageBox.setMessage("Want to save your changes?");
	 	            
	 	            int response = messageBox.open();	 	 	            
	 	            if( response== SWT.YES){	
	 	            	
	 	            	String filePath = directoryPath+"\\"+fileName;
	 	            	boolean isSaved = FileUtils.saveProgress(embeddedWorkbook, filePath, true);
	 	            	
	 	            	if(!isSaved){
	 	            		int styleError = SWT.ICON_ERROR;
	 		        		MessageBox errorMessageBox = Launcher.getInstance().createMessageBox(styleError);
	 		        		errorMessageBox.setMessage("ERROR: Could not save the file!");
	 	            		event.doit = false;
	 	            	}
	 	            	
	 	            	WorkbookUtils.closeEmbeddedWorkbook(embeddedWorkbook, false);
	 	            	Launcher.getInstance().disposeControlSite();
	 	            	Launcher.getInstance().disposeShell();
	 	            	event.doit = true;
	 	            } 
	 	            
	 	            if(response == SWT.NO){
	 	            	Launcher.getInstance().disposeControlSite();
	 	            	Launcher.getInstance().disposeShell();
	 	            	event.doit = true;
	 	            } 
	 	            
	 	            event.doit = false;
	        	}else{
	        		int style = SWT.YES | SWT.NO | SWT.ICON_QUESTION;
	        		MessageBox messageBox = Launcher.getInstance().createMessageBox(style);
	 	            messageBox.setMessage("Do you want to close the application?");
	 	            
	 	            int response = messageBox.open();
	 	            if( response== SWT.YES){
	 	        	    Launcher.getInstance().disposeControlSite();
	 	            	Launcher.getInstance().disposeShell();
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
	        	Launcher.getInstance().setCurrentSelection(RangeUtils.getRangeAddress(rangeAutomation).split(","));
	        	args[0].dispose();
	        	rangeAutomation.dispose();	
	        	
				/*
				 * the second argument is a Worksheet object. Get the name and index of the worksheet.
				 */
	        	OleAutomation worksheetAutomation = args[1].getAutomation();		        
	        	Launcher.getInstance().setActiveWorksheetName(WorksheetUtils.getWorksheetName(worksheetAutomation));
	        	Launcher.getInstance().setActiveWorksheetIndex(WorksheetUtils.getWorksheetIndex(worksheetAutomation));
				args[1].dispose();	
				worksheetAutomation.dispose();
						
				Launcher.getInstance().setFocusToShell();
				Launcher.getInstance().setFocusToShell();

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
				String activeSheetName = WorksheetUtils.getWorksheetName(worksheetAutomation);
				int activeSheetIndex = WorksheetUtils.getWorksheetIndex(worksheetAutomation);
				
				String previousSheetName =Launcher.getInstance().getActiveWorksheetName();
	        	
				// update the information about the active sheet
				Launcher.getInstance().setActiveWorksheetName(activeSheetName);
				Launcher.getInstance().setActiveWorksheetIndex(activeSheetIndex);
				args[0].dispose();
				worksheetAutomation.dispose();
				
				WorksheetAnnotation activeSheetAnnotation = AnnotationHandler.getWorkbookAnnotation()
						.getWorksheetAnnotations().get(activeSheetName);
				
				WorksheetAnnotation previousSheetAnnotation = AnnotationHandler.getWorkbookAnnotation()
						.getWorksheetAnnotations().get(previousSheetName);
										
				if(((activeSheetAnnotation!=null && !activeSheetAnnotation.getAllAnnotations().isEmpty()) || 
						activeSheetName.compareTo(RangeAnnotationsSheet.getName())==0) && previousSheetAnnotation!=null ){ 		
					
						// Keep Redo and Undo list per sheet. Erase when sheet changes
						AnnotationHandler.clearRedoList();
						AnnotationHandler.clearUndoList();
						
						// should not remember selection from previous sheet
						Launcher.getInstance().setCurrentSelection(null);
	        	}
					
					
				// adjust the bar menu according to the properties of the workbook and the active sheet
				BarMenuUtils.adjustBarMenuForWorkbook();		
								
				// return the focus to the embedded excel workbook, if it does not have it already
				if(!Launcher.getInstance().isControlSiteFocusControl())
					Launcher.getInstance().setFocusToControlSite();	
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
	        		Launcher.getInstance().setFocusToControlSite();
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
					Launcher.getInstance().setFocusToControlSite();
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
				
				// warn the user user if there exist an opened file
				// offer them to save their progress
				if(!Launcher.getInstance().isControlSiteNull()  && 
					AnnotationHandler.getWorkbookAnnotation().hashCode()!=
						AnnotationHandler.getOldWorkbookAnnotationHash()){
									
	        		int style = SWT.YES | SWT.NO | SWT.CANCEL | SWT.ICON_WARNING ;
	        		MessageBox messageBox = Launcher.getInstance().createMessageBox(style);
	 	            messageBox.setMessage("Do you want to save the progress on the existing file ?");
	 	            
	 	            int response = messageBox.open();	 	 	            
	 	            if( response== SWT.YES){	
	 	            	OleAutomation embeddedWorkbook = Launcher.getInstance().getEmbeddedWorkbook();
	 	            	String filePath =  Launcher.getInstance().getDirectoryPath()+"\\"+Launcher.getInstance().getFileName();
	 	            	boolean isSaved = FileUtils.saveProgress(embeddedWorkbook, filePath, true);
	 	            	
	 	            	if(!isSaved){
	 	            		int messageStyle = SWT.ICON_ERROR;
	 						MessageBox message = Launcher.getInstance().createMessageBox(messageStyle);
	 						message.setMessage("ERROR: The file could not be saved!");
	 						message.open();
	 	            	}
	 	            } 	 	            
				}
				
				if(!Launcher.getInstance().isControlSiteNull()) {
					OleAutomation embeddedWorkbook  = Launcher.getInstance().getEmbeddedWorkbook();
					WorkbookUtils.closeEmbeddedWorkbook(embeddedWorkbook, false);		
					Launcher.getInstance().setEmbeddedWorkbook(null);
					Launcher.getInstance().disposeControlSite();
				}

				
				// open the files selection window
				FileUtils.fileOpen();
				
				// get the OleAutomation for the embedded workbook
				OleAutomation workbookAutomation = Launcher.getInstance().getEmbeddedWorkbook();
				if (workbookAutomation == null) {
					return; // there is no embedded workbook (file)
				}
								
				// clear all existing annotations in memory structure, 
				// if they exist from the previously opened file 
				AnnotationHandler.getWorkbookAnnotation().removeAllAnnotations();
		
				// create the base in memory structure for storing annotation data
				// retrieve the annotation statuses from previous session
				AnnotationStatusSheet.readAnnotationStatuses(workbookAutomation);
				
				// read the data and re-create the range annotation objects
				RangeAnnotation[] rangeAnnotations = RangeAnnotationsSheet.readRangeAnnotations(workbookAutomation);
				
				if(rangeAnnotations!=null){		
					// update workbook annotation and re-draw all the range annotations  
					AnnotationHandler.recreateRangeAnnotations(workbookAutomation, rangeAnnotations);	
				}
				
				AnnotationHandler.setOldWorkbookAnnotationHash(
						AnnotationHandler.getWorkbookAnnotation().hashCode()); 
				
				// adjust the menu items in the menu bar for the file that was just openned
				BarMenuUtils.adjustBarMenuForOpennedFile();
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
				
				OleAutomation embeddedWorkbook = Launcher.getInstance().getEmbeddedWorkbook();
				String sheetName = Launcher.getInstance().getActiveWorksheetName();
				String fileName = Launcher.getInstance().getFileName();
				String directory = Launcher.getInstance().getDirectoryPath();
				String filePath = directory+"\\"+fileName;
				
				boolean result = FileUtils.saveProgress(embeddedWorkbook, filePath, false);
				if(result){		
					
					AnnotationHandler.clearRedoList();
					AnnotationHandler.clearUndoList();
					
					// to check if the workbook has changed since last save
					int hash = AnnotationHandler.getWorkbookAnnotation().hashCode();
					AnnotationHandler.setOldWorkbookAnnotationHash(hash);
					
					BarMenuUtils.adjustBarMenuForSheet(sheetName);
					
            		int style = SWT.ICON_INFORMATION;
					MessageBox message = Launcher.getInstance().createMessageBox(style);
					message.setMessage("The file was successfully saved.");
					message.open();
				}else{
					int style = SWT.ICON_ERROR;
					MessageBox message = Launcher.getInstance().createMessageBox(style);
					message.setMessage("ERROR: The file could not be saved!");
					message.open();
				}				
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
															
				if(AnnotationHandler.getWorkbookAnnotation().hashCode()!=
						AnnotationHandler.getOldWorkbookAnnotationHash()){
									
	        		int style = SWT.YES | SWT.NO | SWT.CANCEL | SWT.ICON_WARNING ;
	        		MessageBox messageBox = Launcher.getInstance().createMessageBox(style);
	 	            messageBox.setMessage("Want to save your changes?");
	 	            
	 	            int response = messageBox.open();	 	 	            
	 	            if( response== SWT.YES){	
	 	            	OleAutomation embeddedWorkbook = Launcher.getInstance().getEmbeddedWorkbook();
	 	            	String filePath =  Launcher.getInstance().getDirectoryPath()+"\\"+Launcher.getInstance().getFileName();
	 	            	boolean isSaved = FileUtils.saveProgress(embeddedWorkbook, filePath, true);
	 	            	
	 	            	if(!isSaved){
	 	            		int messageStyle = SWT.ICON_ERROR;
	 						MessageBox message = Launcher.getInstance().createMessageBox(messageStyle);
	 						message.setMessage("ERROR: The file could not be saved!");
	 						message.open();
	 	            	}
	 	            } 
	 	            
	 	            if(response == SWT.NO || response == SWT.YES){
	 	            	OleAutomation embeddedWorkbook  = Launcher.getInstance().getEmbeddedWorkbook();
	 					WorkbookUtils.closeEmbeddedWorkbook(embeddedWorkbook, false);
	 					
	 					Launcher.getInstance().setEmbeddedWorkbook(null);
	 					Launcher.getInstance().disposeControlSite();
	 					Color lightGreyShade = new Color (Display.getCurrent(), 247, 247, 247);
	 					Launcher.getInstance().setColorToExcelPanel(lightGreyShade);
	 					
	 					BarMenuUtils.adjustBarMenuForFileClose();
	 	            } 
	 	            
				}else{
					
	        		int style = SWT.YES | SWT.NO | SWT.ICON_QUESTION;
	        		MessageBox messageBox = Launcher.getInstance().createMessageBox(style);
	 	            messageBox.setMessage("Are you sure you want to close the file?");
	 	            
	 	            int response = messageBox.open();
	 	            if( response== SWT.YES){
	 	            	OleAutomation embeddedWorkbook  = Launcher.getInstance().getEmbeddedWorkbook();
	 					WorkbookUtils.closeEmbeddedWorkbook(embeddedWorkbook, false);
	 					
	 					Launcher.getInstance().setEmbeddedWorkbook(null);
	 					Launcher.getInstance().disposeControlSite();
	 					Color lightGreyShade = new Color (Display.getCurrent(), 247, 247, 247);
	 					Launcher.getInstance().setColorToExcelPanel(lightGreyShade);
	 					BarMenuUtils.adjustBarMenuForFileClose();
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
				OleAutomation workbookAutomation = Launcher.getInstance().getEmbeddedWorkbook();
				String directoryPath = Launcher.getInstance().getDirectoryPath();
				String fileName = Launcher.getInstance().getFileName();				
				boolean isSuccess = RangeAnnotationsSheet.exportRangeAnnotationsAsCSV(workbookAutomation, directoryPath, fileName);
				
				if(isSuccess){
					MessageBox messageBox = Launcher.getInstance().createMessageBox(SWT.ICON_INFORMATION);
					messageBox.setText("Information");
		            messageBox.setMessage("The annotation data were successfully exported at:\n"+directoryPath);
		            messageBox.open();
				}else{
					MessageBox messageBox = Launcher.getInstance().createMessageBox(SWT.ICON_ERROR);
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
					MessageBox messageBox = Launcher.getInstance().createMessageBox(SWT.ICON_INFORMATION);
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
				
				if(!Launcher.getInstance().isControlSiteNull() && 
					AnnotationHandler.getWorkbookAnnotation().hashCode()!= AnnotationHandler.getOldWorkbookAnnotationHash()){
								
	        		int style = SWT.YES | SWT.NO | SWT.CANCEL | SWT.ICON_WARNING ;
	        		MessageBox messageBox = Launcher.getInstance().createMessageBox(style);
	 	            messageBox.setMessage("Want to save your changes?");
	 	            
	 	            int response = messageBox.open();	 	 	            
	 	            if( response== SWT.YES){	
	 	            	OleAutomation embeddedWorkbook = Launcher.getInstance().getEmbeddedWorkbook();
	 	            	String filePath =  Launcher.getInstance().getDirectoryPath()+"\\"+Launcher.getInstance().getFileName();
	 	            	boolean isSaved = FileUtils.saveProgress(embeddedWorkbook, filePath, true);
	 	            	
	 	            	if(!isSaved){
	 	            		int messageStyle = SWT.ICON_ERROR;
	 						MessageBox message = Launcher.getInstance().createMessageBox(messageStyle);
	 						message.setMessage("ERROR: The file could not be saved!");
	 						message.open();
	 	            	}else{
	 	            		//String directory = MainWindow.getInstance().getDirectoryPath();
	 	            		//String fileName = MainWindow.getInstance().getFileName();
	 	            		//FileUtils.markFileAsAnnotated(directory, fileName, 1);
	 	            		
		 	            	WorkbookUtils.closeEmbeddedWorkbook(embeddedWorkbook, false);
		 	            	Launcher.getInstance().disposeControlSite();
		 	            	Launcher.getInstance().disposeShell();
	 	            	}
	 	            } 
	 	            
	 	            if(response == SWT.NO){
	 	            	Launcher.getInstance().disposeControlSite();
	 	            	Launcher.getInstance().disposeShell();
	 	            } 
	 	            
	        	}else{
	        		int style = SWT.YES | SWT.NO | SWT.ICON_QUESTION;
	        		MessageBox messageBox = Launcher.getInstance().createMessageBox(style);
	 	            messageBox.setMessage("Do you want to close the application?");
	 	            
	 	            int response = messageBox.open();
	 	            if( response== SWT.YES){
	 	        	    Launcher.getInstance().disposeControlSite();
	 	            	Launcher.getInstance().disposeShell();
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
				String sheetName = Launcher.getInstance().getActiveWorksheetName();
				WorkbookAnnotation workbookAnnotation = AnnotationHandler.getWorkbookAnnotation();
				WorksheetAnnotation  sheetAnnotation = workbookAnnotation.getWorksheetAnnotations().get(sheetName);
				
				if(sheetAnnotation==null){
					System.out.println(sheetName);
					return;
				}
						
				boolean wasUpdated = false;  
				
				if(!sheetAnnotation.isCompleted()){
					if( sheetAnnotation.getAllAnnotations().isEmpty()){
						int style = SWT.ICON_WARNING ;
						MessageBox mb = Launcher.getInstance().createMessageBox(style);
						mb.setMessage("You can not mark this sheet as completed. "
								+ "It does not contain any annotations yet!"); 
						mb.open();
					}else{
						if(sheetAnnotation.getAllAnnotations().size()<4){
							int style = SWT.YES | SWT.NO | SWT.ICON_WARNING ;
							MessageBox mb = Launcher.getInstance().createMessageBox(style);
							mb.setMessage("This sheet contains very few annotations. "
									+ "Do you still want to mark it as \"Completed\" ?"); 
							int option = mb.open();
							
							if(option == SWT.YES){
								sheetAnnotation.setCompleted(true);
								wasUpdated=true;
								
							}
							
						}else{
							sheetAnnotation.setCompleted(true);
							wasUpdated=true;
						}
					}	
				}else{
					sheetAnnotation.setCompleted(false);
					wasUpdated=true;
				}
								
				BarMenuUtils.adjustBarMenuForSheet(sheetName);
				
				if(wasUpdated){
					AnnotationHandler.clearRedoList();
					AnnotationHandler.clearUndoList();
									
					int style = SWT.ICON_INFORMATION;
					MessageBox mb = Launcher.getInstance().createMessageBox(style);
					String value = String.valueOf((sheetAnnotation.isCompleted())).toUpperCase();
					mb.setMessage(" Sheet status was updated to Completed := "+value); 
					mb.open();
				}
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
				String sheetName = Launcher.getInstance().getActiveWorksheetName();
				WorkbookAnnotation workbookAnnotation = AnnotationHandler.getWorkbookAnnotation();
				WorksheetAnnotation  sheetAnnotation = workbookAnnotation.getWorksheetAnnotations().get(sheetName);
				
				if(sheetAnnotation==null){
					System.out.println(sheetName);
					return;
				}
				
				boolean wasUpdated = false;  

				if(!sheetAnnotation.isNotApplicable()){
					
					if(!sheetAnnotation.getAllAnnotations().isEmpty()){
						int style = SWT.YES | SWT.NO | SWT.ICON_WARNING ;
						MessageBox mb = Launcher.getInstance().createMessageBox(style);
						mb.setMessage("Marking this sheet as \"Not Applicable\" will erase all the existing annotations "
								+ "in the sheet. Do you want to proceed ?"); 
						int option = mb.open();
						
						if(option == SWT.YES){
							OleAutomation embeddedWorkbook = Launcher.getInstance().getEmbeddedWorkbook();
							AnnotationHandler.deleteShapeAnnotationsInSheet(embeddedWorkbook, sheetName);
							RangeAnnotationsSheet.deleteRangeAnnotationDataFromSheet(embeddedWorkbook, sheetName, true);
							
							AnnotationHandler.clearRedoList();
							AnnotationHandler.clearUndoList();
							
							workbookAnnotation.removeAllRangeAnnotationsFromSheet(sheetName);
							
							sheetAnnotation.setNotApplicable(true);
							wasUpdated = true;
						}
					}else{
						sheetAnnotation.setNotApplicable(true);
						wasUpdated = true;
					}
					
				}else{
					sheetAnnotation.setNotApplicable(false);
					wasUpdated = true;
				}

				
				BarMenuUtils.adjustBarMenuForSheet(sheetName);
				
				if(wasUpdated){
					AnnotationHandler.clearRedoList();
					AnnotationHandler.clearUndoList();
					
					int style = SWT.ICON_INFORMATION;
					MessageBox mb = Launcher.getInstance().createMessageBox(style);
					String value = String.valueOf((sheetAnnotation.isNotApplicable())).toUpperCase();
					mb.setMessage(" Sheet status was updated to NotApplicable := "+value); 
					mb.open();
				}
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
				
				WorkbookAnnotation wa = AnnotationHandler.getWorkbookAnnotation();		
				boolean wasUpdated = false; 
				
				if(!wa.isCompleted()){

					if( wa.getAllAnnotations().isEmpty()){
						int style = SWT.ICON_WARNING ;
						MessageBox mb = Launcher.getInstance().createMessageBox(style);
						mb.setMessage("You can not mark the file (workbook) as \"Completed\". "
								+ "It does not contain any annotations yet!"); 
						mb.open();
						
					}else{
					
						Collection<WorksheetAnnotation> sheetAnnotations 
											= wa.getWorksheetAnnotations().values();
						Iterator<WorksheetAnnotation> itr = sheetAnnotations.iterator();
						
						boolean hasPendingSheets =  false;
						while (itr.hasNext()) {
							WorksheetAnnotation sa = itr.next();
							
							if(!sa.isCompleted() && !sa.isNotApplicable()){
								if(sa.getAllAnnotations().isEmpty() || sa.getAllAnnotations().size()<4){
									
									hasPendingSheets = true; 
									
									OleAutomation embeddedWorkbook = Launcher.getInstance().getEmbeddedWorkbook();
									OleAutomation worksheetAutomation = 
											WorkbookUtils.getWorksheetAutomationByName(embeddedWorkbook, sa.getSheetName());
									WorksheetUtils.makeWorksheetActive(worksheetAutomation);
									
									int style = SWT.YES | SWT.NO | SWT.ICON_WARNING ;
									MessageBox mb = Launcher.getInstance().createMessageBox(style);
									mb.setMessage("The \""+sa.getSheetName()+"\" does not contain any or "
											+ "contains very few annotations. "
											+ "\nDo you still want to mark this file (workbook) as \"Completed\" ?"); 
									int option = mb.open();
									
									if(option == SWT.YES){
										wasUpdated=true;
										wa.setCompleted(true);
									}
									
									if(option == SWT.NO){
										wasUpdated = false;
										AnnotationHandler.clearRedoList();
										AnnotationHandler.clearUndoList();
										break;
									}		
									
								}
							}
						}
						
						if(!hasPendingSheets ){
							wa.setCompleted(true);
							wasUpdated=true;
						}
					}
				}else{
					wa.setCompleted(false);
					wasUpdated=true;
				}
				
				BarMenuUtils.adjustBarMenuForWorkbook();
				
				if(wasUpdated){			
					AnnotationHandler.clearRedoList();
					AnnotationHandler.clearUndoList();
					
					BarMenuUtils.adjustBarMenuForWorkbook();
					
					int style = SWT.ICON_INFORMATION;
					MessageBox mb = Launcher.getInstance().createMessageBox(style);
					String value = String.valueOf((wa.isCompleted())).toUpperCase();
					mb.setMessage("File (Workbook) status was updated to Completed := "+value); 
					mb.open();
				}
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
				boolean wasUpdated = false;
				
				if(!workbookAnnotation.isNotApplicable()){
					
					if(!workbookAnnotation.getAllAnnotations().isEmpty()){
						int style = SWT.YES | SWT.NO | SWT.ICON_WARNING ;
						MessageBox mb = Launcher.getInstance().createMessageBox(style);
						mb.setMessage("Marking this file (workbook) as \"Not Applicable\" "
								+ "will erase all the existing annotations. "
								+ "Do you want to proceed ?"); 
						int option = mb.open();
						
						if(option == SWT.YES){
							OleAutomation embeddedWorkbook = Launcher.getInstance().getEmbeddedWorkbook();
							AnnotationHandler.deleteAllShapeAnnotations(embeddedWorkbook);
							RangeAnnotationsSheet.deleteAllRangeAnnotationData(embeddedWorkbook);
							
							AnnotationHandler.clearRedoList();
							AnnotationHandler.clearUndoList();
							AnnotationHandler.getWorkbookAnnotation().removeAllAnnotations();
							
							workbookAnnotation.setNotApplicable(true);
							wasUpdated = true;
						}
					}else{
						workbookAnnotation.setNotApplicable(true);
						wasUpdated = true;
					}
					
				}else{
					workbookAnnotation.setNotApplicable(false);
					wasUpdated = true;
				}
				
				BarMenuUtils.adjustBarMenuForWorkbook();
				
				if(wasUpdated){
					
					int style = SWT.ICON_INFORMATION;
					MessageBox mb = Launcher.getInstance().createMessageBox(style);
					String value = String.valueOf((workbookAnnotation.isNotApplicable())).toUpperCase();
					mb.setMessage("File (Workbook) status was updated to NotApplicable := "+value); 
					mb.open();
				}
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
				 
				 OleAutomation workbookAutomation = Launcher.getInstance().getEmbeddedWorkbook();
				 
				 String sheetName = Launcher.getInstance().getActiveWorksheetName();
				 int sheetIndex = Launcher.getInstance().getActiveWorksheetIndex();
				 String[] currentSelection = Launcher.getInstance().getCurrentSelection();
				 
				 try{
					 AnnotationHandler.annotate(workbookAutomation, sheetName, sheetIndex,   
				 		 currentSelection, annotationClass);				 
				 }catch (Exception ex){
					 logger.error("Generic exception on create new annotation", ex);
				 }
				 
				 // if the sheet was empty, had no annotations, 
				 // the menu needs to be updated
				 BarMenuUtils.adjustBarMenuForSheet(sheetName);
				 
				 if(Launcher.getInstance().isControlSiteFocusControl())
					 	Launcher.getInstance().setFocusToShell();			 
			}
		};
	}
	
	/**
	 * 
	 * @return
	 */
	protected static SelectionListener createUndoLastAnnotationSelectionListener(){
		return new SelectionAdapter() {
			@Override
			public void widgetSelected(SelectionEvent e) {
		
				RangeAnnotation  ra = AnnotationHandler.getLastFromUndoList();	
				if(ra==null)
					return;
				
				OleAutomation workbookAutomation = Launcher.getInstance().getEmbeddedWorkbook(); 
				OleAutomation sheetAutomation = 
						WorkbookUtils.getWorksheetAutomationByName(workbookAutomation, ra.getSheetName());
				
				WorksheetUtils.unprotectWorksheet(sheetAutomation);		
				boolean isSuccess = AnnotationHandler.deleteAnnotationFromSheet(sheetAutomation, ra);
				WorksheetUtils.protectWorksheet(sheetAutomation);
				sheetAutomation.dispose();
				
				if(isSuccess){
					AnnotationHandler.removeLastFromUndoList();
					AnnotationHandler.addToRedoList(ra);
						
					RangeAnnotationsSheet.deleteRangeAnnotationData(workbookAutomation, ra, true);
					
					AnnotationHandler.getWorkbookAnnotation().removeRangeAnnotation(ra);
						
					Launcher.getInstance().setActiveWorksheetIndex(ra.getSheetIndex());
					Launcher.getInstance().setActiveWorksheetName(ra.getSheetName());
					Launcher.getInstance().setCurrentSelection(new String[]{ra.getRangeAddress()});
									
				}else{
					MessageBox messageBox = Launcher.getInstance().createMessageBox(SWT.ICON_ERROR);
	 	            messageBox.setMessage("Could not undo the last range annotation!!!");
	 	            messageBox.open();
				}
					
				BarMenuUtils.adjustBarMenuForSheet(ra.getSheetName());
			}
		};
	}	
	
	/**
	 * 
	 * @return
	 */
	protected static SelectionListener createRedoLastAnnotationSelectionListener(){
		return new SelectionAdapter() {
			@Override
			public void widgetSelected(SelectionEvent e) {
				
				RangeAnnotation  ra = AnnotationHandler.getLastFromRedoList();		
				if(ra==null){
					return;
				}
				
				OleAutomation workbookAutomation = Launcher.getInstance().getEmbeddedWorkbook(); 
				OleAutomation worksheetAutomation = 
						WorkbookUtils.getWorksheetAutomationByName(workbookAutomation, ra.getSheetName());
				
				WorksheetUtils.unprotectWorksheet(worksheetAutomation);		
				
				
				Boolean result = false;
				try{
					result = AnnotationHandler.drawRangeAnnotation(workbookAutomation, ra, true);
				}catch (Exception ex){			
					logger.error("Generic exception on redo last annotation", ex);
				}
				
				WorksheetUtils.protectWorksheet(worksheetAutomation);
				worksheetAutomation.dispose();
				
				AnnotationHandler.removeLastFromRedoList();
				if(!result){
					AnnotationHandler.getWorkbookAnnotation().removeRangeAnnotation(ra);
					BarMenuUtils.adjustBarMenuForSheet(ra.getSheetName());	
					return;
				}
				
				AnnotationHandler.addToUndoList(ra);	
				
				RangeAnnotationsSheet.saveRangeAnnotationData(workbookAutomation, ra);
				AnnotationHandler.getWorkbookAnnotation().addRangeAnnotation(ra);
				
				BarMenuUtils.adjustBarMenuForSheet(ra.getSheetName());			
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
				OleAutomation workbookAutomation = Launcher.getInstance().getEmbeddedWorkbook();	
				AnnotationHandler.setVisilityForAllAnnotations(workbookAutomation, false);
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
				OleAutomation workbookAutomation = Launcher.getInstance().getEmbeddedWorkbook();
				String sheetName = Launcher.getInstance().getActiveWorksheetName();
				AnnotationHandler.setVisibilityForAnnotationsInSheet(workbookAutomation, sheetName, false);
			}
		};
	}

	/**
	 * 
	 * @return
	 */
	protected static SelectionListener createShowAllAnnotationsSelectionListener(){
		return new SelectionAdapter() {
			
			@Override
			public void widgetSelected(SelectionEvent e) {
				OleAutomation workbookAutomation = Launcher.getInstance().getEmbeddedWorkbook();
				AnnotationHandler.setVisilityForAllAnnotations(workbookAutomation, true);
			}			
		};
	}
	
	/**
	 * 
	 * @return
	 */
	protected static SelectionListener createShowInSheetAnnotationsSelectionListener(){
		return new SelectionAdapter() {
			
			@Override
			public void widgetSelected(SelectionEvent e) {
				OleAutomation workbookAutomation = Launcher.getInstance().getEmbeddedWorkbook();
				String sheetName = Launcher.getInstance().getActiveWorksheetName();
				AnnotationHandler.setVisibilityForAnnotationsInSheet(workbookAutomation, sheetName, true);
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
				OleAutomation workbookAutomation = Launcher.getInstance().getEmbeddedWorkbook();	
				String sheetName = Launcher.getInstance().getActiveWorksheetName();
				
				AnnotationHandler.deleteAllShapeAnnotations(workbookAutomation);
				
				WorkbookAnnotation workbookAnnotation = AnnotationHandler.getWorkbookAnnotation();
				workbookAnnotation.removeAllAnnotations();
				
				RangeAnnotationsSheet.deleteAllRangeAnnotationData(workbookAutomation);
				
				AnnotationHandler.clearRedoList();
				AnnotationHandler.clearUndoList();
				
				BarMenuUtils.adjustBarMenuForSheet(sheetName);
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

				OleAutomation workbookAutomation = Launcher.getInstance().getEmbeddedWorkbook();
				String sheetName = Launcher.getInstance().getActiveWorksheetName();
				
				AnnotationHandler.deleteShapeAnnotationsInSheet(workbookAutomation, sheetName);
				
				WorkbookAnnotation workbookAnnotation = AnnotationHandler.getWorkbookAnnotation();
				workbookAnnotation.removeAllRangeAnnotationsFromSheet(sheetName);
				
				RangeAnnotationsSheet.deleteRangeAnnotationDataFromSheet(workbookAutomation, sheetName, true);		
				
				AnnotationHandler.clearRedoList();
				AnnotationHandler.clearUndoList();
				
				BarMenuUtils.adjustBarMenuForSheet(sheetName);
			}
		};
	}
}
