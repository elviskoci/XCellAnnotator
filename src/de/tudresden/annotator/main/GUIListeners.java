/**
 * 
 */
package de.tudresden.annotator.main;

import org.eclipse.swt.SWT;
import org.eclipse.swt.ole.win32.OleAutomation;
import org.eclipse.swt.ole.win32.OleEvent;
import org.eclipse.swt.ole.win32.OleListener;
import org.eclipse.swt.ole.win32.Variant;
import org.eclipse.swt.widgets.Event;
import org.eclipse.swt.widgets.Listener;
import org.eclipse.swt.widgets.MessageBox;

import de.tudresden.annotator.oleutils.RangeUtils;
import de.tudresden.annotator.oleutils.WorkbookUtils;
import de.tudresden.annotator.oleutils.WorksheetUtils;

/**
 * @author Elvis Koci
 */
public class GUIListeners {
	
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
	 	            	boolean isSaved = FileManager.saveProgress(embeddedWorkbook, filePath);
	 	            	
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
	 * Create a SheetSelection event listener
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
	 * Create a SheetActivate event listener
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
				MainWindow.getInstance().setActiveWorksheetName(WorksheetUtils.getWorksheetName(worksheetAutomation));
				MainWindow.getInstance().setActiveWorksheetIndex(WorksheetUtils.getWorksheetIndex(worksheetAutomation));
				args[0].dispose();
				worksheetAutomation.dispose();
				
				MainWindow.getInstance().setFocusToControlSite();
	        }
	    };	       
	    return listener;
	}
	
	
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
	
	
	protected static Listener createMouseWheelEventListener(){		
		return new Listener() {
			@Override
			public void handleEvent(Event e) {
				MainWindow.getInstance().setFocusToControlSite();
			}		
		};
	}
}
