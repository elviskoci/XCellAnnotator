/**
 * 
 */
package de.tudresden.annotator.main;

import java.io.File;

import org.eclipse.swt.SWT;
import org.eclipse.swt.SWTError;
import org.eclipse.swt.layout.FillLayout;
import org.eclipse.swt.ole.win32.OLE;
import org.eclipse.swt.ole.win32.OleAutomation;
import org.eclipse.swt.ole.win32.OleControlSite;
import org.eclipse.swt.ole.win32.OleEvent;
import org.eclipse.swt.ole.win32.OleFrame;
import org.eclipse.swt.ole.win32.OleListener;
import org.eclipse.swt.ole.win32.Variant;
import org.eclipse.swt.widgets.Display;
import org.eclipse.swt.widgets.FileDialog;
import org.eclipse.swt.widgets.MessageBox;
import org.eclipse.swt.widgets.Shell;

/**
 * @author Elvis
 *
 */
public class MainWindow {
	
	// GUID Event Sink
	private static final String IID_AppEvents = "{00024413-0000-0000-C000-000000000046}";
	// Event ID
	private static final int SheetSelectionChange   = 0x00000616;
		
	private final Display display = new Display();;
	private final Shell shell = new Shell(display);
	private OleFrame oleFrame;
	private OleControlSite controlSite;
	
	private String currentSelection[];
	private String activeWorksheetName;
	private long activeWorksheetIndex;
	
	private static MainWindow instance = null;
	private MainWindow() {
		//display = new Display();
		//shell = new Shell(display);
    }
  
	public static MainWindow getInstance() {
		if(instance == null) {
			instance = new MainWindow();
		}
		return instance;  
	}
	
	/**
	 * Create the window that will serve as the main Graphical User Interface (GUI) for user interaction
	 * @param shell
	 */
	private void buildGUIWindow(Shell shell){
		shell.setText("Annotator");
	    shell.setLayout(new FillLayout());
	    shell.setSize(1200, 600);
		
	    oleFrame = new OleFrame(shell, SWT.NONE);
	    
	    BarMenu  oleFrameMenuBar = new BarMenu(oleFrame.getShell());
	    oleFrame.setFileMenus(oleFrameMenuBar.getMenuItems());
	}
	
	private void adjustSpreadsheetDisplay(OleAutomation application){
		
		if(controlSite==null)
			return;
		 
	    // Create custom Cell commandbar	
//		OleInterfaceModifier.createCustomCellCommandBar(application);
	    
	    // undoChangesToCellCommandBar(application);
//	    controlSite.doVerb(OLE.OLEIVERB_INPLACEACTIVATE);	
	    
		// minimize ribbon	
		// TODO: Individual CommandBars
	    OleInterfaceModifier.hideRibbon(application);	
	    
	    // hide menu on right click of user at a worksheet tab
//	    CommandBarsHelper.setVisibilityForCommandBar(application, "Ply", false);
	    CommandBarsHelper.setEnabledForCommandBar(application, "Ply", false);
	    
	    // hide menu on right click of user on a cell
//	    CommandBarsHelper.setVisibilityForCommandBar(application, "Cell", false);
	    CommandBarsHelper.setEnabledForCommandBar(application, "Cell", false);
	    
	    // get active workbook, the one that is loaded by this application
	    OleAutomation activeWorkbook =  OleInterfaceModifier.getActiveWorkbook(application);	    
	    // protect the structure of the active workbook
	    if(!OleInterfaceModifier.protectWorkbook(activeWorkbook))
	    	System.out.println("\nERROR: Unable to protect active workbook!");
	    
	    // protect all individual worksheets
	    if(!OleInterfaceModifier.protectAllWorksheets(activeWorkbook))
	    	System.out.println("\nERROR: Unable to protect the worksheets that are part of the active workbook!");
	   
	    activeWorkbook.dispose();
	}

	/**
	 * Open an excel file for annotation
	 */
	 public void fileOpen(){
		
		// Select the excel file
		Shell shell = getOleFrame().getShell();
		FileDialog dialog = new FileDialog(shell, SWT.OPEN);
		String fileName = dialog.open();
		
		// if no file was selected, return
		if (fileName == null) return;
		
		// dispose OleControlSite if it is not null. 
		disposeControlSite();
				
	    if (getControlSite() == null) {
			int index = fileName.lastIndexOf('.');
			if (index != -1) {
				String fileExtension = fileName.substring(index + 1);
				if (fileExtension.equalsIgnoreCase("xls") || fileExtension.equalsIgnoreCase("xlsx")) {	
					
					// create a new control site to open the file with Excel
					try {		    	
				        
						// create new OLE control site using the specified excel file
						File excelFile = new File(fileName);
				        setControlSite(new OleControlSite(oleFrame, SWT.NONE, excelFile));
				        
				        // add sheet selection event listener 
					    OleAutomation application = OleInterfaceModifier.getApplicationAutomation(getControlSite());
				        OleListener listener = createSheetSelectionEventListener(application);
				        getControlSite().addEventListener(application, IID_AppEvents, SheetSelectionChange, listener);
				       
				        
				        // activate and display excel workbook
				        getControlSite().doVerb(OLE.OLEIVERB_INPLACEACTIVATE);	
				        
				        // adjust the excel application user interface for the annotation task
				        adjustSpreadsheetDisplay(application);
				        application.dispose();
				        
				    } catch (SWTError e) {
				        e.printStackTrace();
				        System.out.println("Unable to open ActiveX Control");
				        return;
				    }	    	  
				   
				}else{
					MessageBox msgbox = new MessageBox(shell,SWT.ICON_ERROR);
					msgbox.setMessage("The selected file format is not recognized: ."+fileExtension);
					msgbox.open();
				}
			}
	    }
	}
	
	 
	/**
	 * Create a SheetSelection event listener
	 * @param application
	 * @return
	 */
	private OleListener createSheetSelectionEventListener(OleAutomation application){
		
		OleListener listener = new OleListener() {
	        public void handleEvent (OleEvent e) {
	        	
	        	Variant[] args = e.arguments;
	        	
	            /*
	             * the first argument is a Range object. get the number and range of selected areas 
	             */
	        	OleAutomation rangeAutomation = args[0].getAutomation();
				
	        	int[] addressIds = rangeAutomation.getIDsOfNames(new String[]{"Address"}); 
				Variant addressVariant = rangeAutomation.getProperty(addressIds[0]);	
//					System.out.print("The selection has changed to: {"+addressVariant.getString()+"}. ");
				setCurrentSelection(addressVariant.getString().split(","));
				addressVariant.dispose();
				
				int[] areasIds = rangeAutomation.getIDsOfNames(new String[]{"Areas"}); 
				Variant areasVariant = rangeAutomation.getProperty(areasIds[0]);								
				OleAutomation areasAutomation = areasVariant.getAutomation();
				areasVariant.dispose();
				
				int[] countId = areasAutomation.getIDsOfNames(new String[]{"Count"});									
				Variant  countVariant = areasAutomation.getProperty(countId[0]);
//					System.out.println("It includes "+countVariant.getString()+" area/s.");
				countVariant.dispose();
				
				args[0].dispose();
				rangeAutomation.dispose();
							
				/*
				 * the second argument is a Worksheet object. get the name and index of the worksheet 	
				 */
				OleAutomation worksheetAutomation = args[1].getAutomation();
				
				int[] nameIds = worksheetAutomation.getIDsOfNames(new String[]{"Name"}); 
				Variant nameVariant = worksheetAutomation.getProperty(nameIds[0]);	
//					System.out.print("Selection has occured at worksheet \""+nameVariant.getString()+"\", ");
				setActiveWorksheetName(nameVariant.getString());
				nameVariant.dispose();
				
				int[] indexIds = worksheetAutomation.getIDsOfNames(new String[]{"Index"}); 
				Variant indexVariant = worksheetAutomation.getProperty(indexIds[0]);	
//					System.out.println("which has indexNo "+indexVariant.getString()+".\n");
				setActiveWorksheetIndex(indexVariant.getLong());
				indexVariant.dispose();
				
				args[1].dispose();
				worksheetAutomation.dispose();
	        }
	    };	       
	    return listener;
	}
	 
    /**
	 * Create message box using the "main" window (this class) shell 
	 * @param style 
	 * @return
	 */
	public MessageBox createMessageBox(int style){
		return new MessageBox(shell,style);
	}
			
	/**
	 * Dispose control site 
	 */
	public void disposeControlSite() {
		if (controlSite != null){
			controlSite.dispose();
		}
		controlSite = null;
	}
	
	/**
	 * Dispose shell
	 */
	public void disposeShell() {
		if (shell != null){
			shell.dispose();
		}
	}
	
	/**
	 * Get OleFrame
	 * @return
	 */
	public OleFrame getOleFrame() {
		return oleFrame;
	}
		
	
	/**
	 * Set OleFrame
	 * 
	 * @param oleFrame
	 */
	public void setOleFrame(OleFrame oleFrame) {
		this.oleFrame = oleFrame;
	}
	
	/**
	 * Get OleControlSite
	 * @return
	 */
	public OleControlSite getControlSite() {
		return controlSite;
	}
	
	/**
	 * Set OleControlSite
	 * @param controlSite
	 */
	public void setControlSite(OleControlSite controlSite) {
		this.controlSite = controlSite;
	}
	
	/**
	 * @return the currentSelection
	 */
	protected String[] getCurrentSelection() {
		return currentSelection;
	}

	/**
	 * @param currentSelection the currentSelection to set
	 */
	protected void setCurrentSelection(String[] currentSelection) {
		this.currentSelection = currentSelection;
	}

	/**
	 * @return the activeWorksheetName
	 */
	protected String getActiveWorksheetName() {
		return activeWorksheetName;
	}

	/**
	 * @param activeWorksheetName the activeWorksheetName to set
	 */
	protected void setActiveWorksheetName(String activeWorksheetName) {
		this.activeWorksheetName = activeWorksheetName;
	}

	/**
	 * @return the activeWorksheetIndex
	 */
	protected long getActiveWorksheetIndex() {
		return activeWorksheetIndex;
	}

	/**
	 * @param activeWorksheetIndex the activeWorksheetIndex to set
	 */
	protected void setActiveWorksheetIndex(long activeWorksheetIndex) {
		this.activeWorksheetIndex = activeWorksheetIndex;
	}

	/**
	 * @return the display
	 */
	public Display getDisplay() {
		return display;
	}

	/**
	 * @return the shell
	 */
	public Shell getShell() {
		return shell;
	}

	/**
	 * @param args
	 */
	public static void main(String[] args) {

		MainWindow GUI = MainWindow.getInstance(); 
		
	    GUI.buildGUIWindow(GUI.getShell());

  		GUI.getShell().open();
  		
  	    while (!GUI.getShell().isDisposed ()) {
  	        if (!GUI.getDisplay().readAndDispatch ()) GUI.getDisplay().sleep();
  	    }
  	    
	    GUI.getDisplay().dispose();
	}		
}
