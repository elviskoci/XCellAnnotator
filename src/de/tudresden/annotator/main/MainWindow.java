/**
 * 
 */
package de.tudresden.annotator.main;


import java.io.File;

import org.eclipse.swt.SWT;
import org.eclipse.swt.SWTError;
import org.eclipse.swt.custom.SashForm;
import org.eclipse.swt.events.FocusEvent;
import org.eclipse.swt.events.FocusListener;
import org.eclipse.swt.events.KeyEvent;
import org.eclipse.swt.events.KeyListener;
import org.eclipse.swt.graphics.Color;
import org.eclipse.swt.layout.FillLayout;
import org.eclipse.swt.ole.win32.OLE;
import org.eclipse.swt.ole.win32.OleAutomation;
import org.eclipse.swt.ole.win32.OleControlSite;
import org.eclipse.swt.ole.win32.OleEvent;
import org.eclipse.swt.ole.win32.OleFrame;
import org.eclipse.swt.ole.win32.OleListener;
import org.eclipse.swt.ole.win32.Variant;
import org.eclipse.swt.widgets.Composite;
import org.eclipse.swt.widgets.Display;
import org.eclipse.swt.widgets.Event;
import org.eclipse.swt.widgets.FileDialog;
import org.eclipse.swt.widgets.Label;
import org.eclipse.swt.widgets.Listener;
import org.eclipse.swt.widgets.MessageBox;
import org.eclipse.swt.widgets.Shell;

import de.tudresden.annotator.oleutils.ApplicationUtils;
import de.tudresden.annotator.oleutils.CommandBarUtils;
import de.tudresden.annotator.oleutils.RangeUtils;
import de.tudresden.annotator.oleutils.WorkbookUtils;
import de.tudresden.annotator.oleutils.WorksheetUtils;

/**
 * @author Elvis
 *
 */
public class MainWindow {
	
	// GUID Event Sink
	private final String IID_AppEvents = "{00024413-0000-0000-C000-000000000046}";
	// Event IDs
	private final int SheetSelectionChange = 0x00000616;
	private final int SheetActivate        = 0x00000619;
		
	private final Display display = new Display();;
	private final Shell shell = new Shell(display);
	
	private OleFrame oleFrame;
	private OleControlSite controlSite;
	
	private OleAutomation excelApplication;
	
	private OleAutomation embeddedWorkbook;
	private String embeddedWorkbookName;
	private String directoryPath;
	private String fileName;
	
	private String activeWorksheetName;
	private int activeWorksheetIndex;
	
	private String currentSelection[];
		
	private static MainWindow instance = null;
	private MainWindow(){}
	
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
		
		Color white = new Color (Display.getCurrent(), 255, 255, 255);
		Color lightGreyShade = new Color (Display.getCurrent(), 247, 247, 247);
		// Color black = new Color (Display.getCurrent(), 0, 0, 0);
		// Color lightBlue = new Color(Display.getCurrent(),229, 248, 255); 
		
		// Shell properties
		shell.setText("Annotator");
	    shell.setLayout(new FillLayout());
	    shell.setSize(1500, 800);
	    
	    shell.addListener(SWT.Close, new Listener()
	    {
	        public void handleEvent(Event event)
	        {	
	        	if(controlSite!=null && controlSite.isDirty()){
	        		int style = SWT.YES | SWT.NO | SWT.CANCEL | SWT.ICON_WARNING ;
	        		MessageBox messageBox = new MessageBox(shell, style);
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
	 	            	disposeControlSite();
	 	            	disposeShell();
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
	    });
	    
	    // Split shell in two horizontal panels 
	    SashForm mainSash = new SashForm(shell, SWT.HORIZONTAL);
		mainSash.setLayout(new FillLayout());
		//mainSash.setBackground(greyShade);
	
		// the left panel will contain the folder explorer. That is a tree structure of files and folders.
		Composite leftPanel = new Composite(mainSash, SWT.BORDER );
		leftPanel.setLayout(new FillLayout());
	
		Label leftPanelLabel = new Label(leftPanel, SWT.NONE);
		leftPanelLabel.setText("Folder Explorer");
		leftPanelLabel.setBackground(white);
		
		// the right panel will be subdivided into two more panels
		Composite rightPanel = new Composite(mainSash, SWT.BORDER);
		rightPanel.setLayout(new FillLayout());
		
		mainSash.setWeights(new int[] {10,90});
			
		// Sub split the right panel 
	    SashForm rightSash = new SashForm(rightPanel, SWT.VERTICAL);
		rightSash.setLayout(new FillLayout());
		
		// Create the panel that will embed the excel application
		Composite excelPanel =  new Composite(rightSash, SWT.BORDER );
		excelPanel.setLayout(new FillLayout());
		
		setOleFrame(new OleFrame(excelPanel, SWT.NONE));
		getOleFrame().setBackground(lightGreyShade);
		
		// Create the panel that will display the applied annotations for the current file
		Composite annotationsPanel =  new Composite(rightSash, SWT.BORDER );
		annotationsPanel.setLayout(new FillLayout());
		
		Label bottomPanelLabel = new Label(annotationsPanel, SWT.NONE);
		bottomPanelLabel.setText("Annotations Panel");
		bottomPanelLabel.setBackground(white);
		
		rightSash.setWeights(new int[] {80,20});
		
		// add a bar menu
	    BarMenu  oleFrameMenuBar = new BarMenu(getOleFrame().getShell());
	    getOleFrame().setFileMenus(oleFrameMenuBar.getMenuItems());		    
	}
	
	private void setUpWorkbookDisplay(){
		
		if(getControlSite()==null){
			System.out.println("Control Site is null! Cannot proceed with the display set up.");
			System.exit(1);
		}
	
		// get excel application as OLE automation object
	    OleAutomation application = ApplicationUtils.getApplicationAutomation(getControlSite());
	    // TODO: suppress alerts
	        
	    // add event listeners
	    OleListener sheetSelectionListener = createSheetSelectionEventListener(application);
        getControlSite().addEventListener(application, IID_AppEvents, SheetSelectionChange, sheetSelectionListener);
        
        OleListener sheetActivationlistener = createSheetActivationEventListener(application);
        getControlSite().addEventListener(application, IID_AppEvents, SheetActivate, sheetActivationlistener);
            
		// minimize ribbon.	TODO: Try hiding individual CommandBars
	    ApplicationUtils.hideRibbon(application);	
	    
	    // hide menu on right click of user at a worksheet tab
	    CommandBarUtils.setEnabledForCommandBar(application, "Ply", true);
	    
	    // hide menu on right click of user on a cell
	    CommandBarUtils.setEnabledForCommandBar(application, "Cell", false);
	    
	    // get active workbook, the one that is embedded in this application
	    OleAutomation workbook = ApplicationUtils.getActiveWorkbookAutomation(application);
	    setEmbeddedWorkbook(workbook);
	    
	    // get the name of workbook for future reference. The name of the workbook might be different from the excel file name. 
	    String workbookName = WorkbookUtils.getWorkbookName(workbook);
	    
	    // get the active worksheet automation, i.e. the one that is on top of the other worksheet
	    OleAutomation worksheet = WorkbookUtils.getActiveWorksheetAutomation(workbook);
	    
	    // get and store the name and index for the worksheet that is "active"
	    String sheetName = WorksheetUtils.getWorksheetName(worksheet);
	    setActiveWorksheetName(sheetName);
	    int sheetIndex = WorksheetUtils.getWorksheetIndex(worksheet);
	    setActiveWorksheetIndex(sheetIndex);
	    
	    setEmbeddedWorkbookName(workbookName);
	    setExcelApplication(application);
	    worksheet.dispose();
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
				if (fileExtension.equalsIgnoreCase("xls") || fileExtension.equalsIgnoreCase("xlsx") || fileExtension.equalsIgnoreCase("xlsm")) { // including macro enabled ?	
					
					try {		    	
				        
						File excelFile = new File(fileName);
						setDirectoryPath(excelFile.getParent());
						setFileName(excelFile.getName());
						
						// create new OLE control site to open the specified excel file
				        setControlSite(new OleControlSite(getOleFrame(), SWT.NONE, "Excel.Sheet" ,excelFile));
				        
				        // activate and display excel workbook
				        getControlSite().doVerb(OLE.OLEIVERB_INPLACEACTIVATE);	
				        
				        // set up the excel application user interface for the annotation task
				        setUpWorkbookDisplay();
				        
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
	             * the first argument is a Range object. Get the number and range of selected areas 
	             */	        	
	        	OleAutomation rangeAutomation = args[0].getAutomation();
	        	setCurrentSelection(RangeUtils.getRangeAddress(rangeAutomation).split(","));
	        	args[0].dispose();
	        	rangeAutomation.dispose();	
	        	
				/*
				 * the second argument is a Worksheet object. Get the name and index of the worksheet.
				 */
	        	OleAutomation worksheetAutomation = args[1].getAutomation();		        
				setActiveWorksheetName(WorksheetUtils.getWorksheetName(worksheetAutomation));
				setActiveWorksheetIndex(WorksheetUtils.getWorksheetIndex(worksheetAutomation));
				args[1].dispose();	
				worksheetAutomation.dispose();
						
				MainWindow.getInstance().getShell().setFocus();
				MainWindow.getInstance().getShell().setFocus();
	        }
	    };	       
	    return listener;
	}
	
	
	/**
	 * Create a SheetActivate event listener
	 * @param application
	 * @return
	 */
	private OleListener createSheetActivationEventListener(OleAutomation application){
		
		OleListener listener = new OleListener() {
	        public void handleEvent (OleEvent e) {
	        	
	        	Variant[] args = e.arguments;
	        	
	        	/*
	             * This event returns only one argument, a Worksheet object. Get the name and index of the activated worksheet.
	             */ 					
				OleAutomation worksheetAutomation = args[0].getAutomation();        
				setActiveWorksheetName(WorksheetUtils.getWorksheetName(worksheetAutomation));
				setActiveWorksheetIndex(WorksheetUtils.getWorksheetIndex(worksheetAutomation));
				args[0].dispose();
				worksheetAutomation.dispose();
				
				MainWindow.getInstance().getControlSite().setFocus();
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
	protected void disposeControlSite() {
		if (controlSite != null){
			controlSite.dispose();
		}
		controlSite = null;
	}
	
	/**
	 * Dispose shell
	 */
	protected void disposeShell() {
		if (shell != null){
			shell.dispose();
		}
	}
	
	/**
	 * @return the display
	 */
	protected Display getDisplay() {
		return display;
	}

	/**
	 * @return the shell
	 */
	protected Shell getShell() {
		return shell;
	}

	/**
	 * Get OleFrame
	 * @return
	 */
	protected OleFrame getOleFrame() {
		return oleFrame;
	}
		
	
	/**
	 * Set OleFrame
	 * 
	 * @param oleFrame
	 */
	protected void setOleFrame(OleFrame oleFrame) {
		this.oleFrame = oleFrame;
	}
	
	/**
	 * Get OleControlSite
	 * @return
	 */
	protected OleControlSite getControlSite() {
		return controlSite;
	}
	
	/**
	 * Set OleControlSite
	 * @param controlSite
	 */
	protected void setControlSite(OleControlSite controlSite) {
		this.controlSite = controlSite;
	}
	
	/**
	 * @return the excelApplication
	 */
	protected OleAutomation getExcelApplication() {
		return excelApplication;
	}

	/**
	 * @param excelApplication the excelApplication to set
	 */
	protected void setExcelApplication(OleAutomation excelApplication) {
		this.excelApplication = excelApplication;
	}

	/**
	 * @return the embeddedWorkbook
	 */
	protected OleAutomation getEmbeddedWorkbook() {
		return embeddedWorkbook;
	}

	/**
	 * @param embeddedWorkbook the embeddedWorkbook to set
	 */
	protected void setEmbeddedWorkbook(OleAutomation embeddedWorkbook) {
		this.embeddedWorkbook = embeddedWorkbook;
	}
	
	/**
	 * @return the activeWorkbookName
	 */
	protected String getEmbeddedWorkbookName() {
		return embeddedWorkbookName;
	}
	
	/**
	 * @param activeWorkbookName the activeWorkbookName to set
	 */
	protected void setEmbeddedWorkbookName(String activeWorkbookName) {
		this.embeddedWorkbookName = activeWorkbookName;
	}
		
	/**
	 * @return the filePath
	 */
	protected String getDirectoryPath() {
		return directoryPath;
	}

	/**
	 * @param filePath the filePath to set
	 */
	protected void setDirectoryPath(String filePath) {
		this.directoryPath = filePath;
	}

	/**
	 * @return the fileName
	 */
	protected String getFileName() {
		return fileName;
	}

	/**
	 * @param fileName the fileName to set
	 */
	protected void setFileName(String fileName) {
		this.fileName = fileName;
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
	protected int getActiveWorksheetIndex() {
		return activeWorksheetIndex;
	}

	/**
	 * @param activeWorksheetIndex the activeWorksheetIndex to set
	 */
	protected void setActiveWorksheetIndex(int activeWorksheetIndex) {
		this.activeWorksheetIndex = activeWorksheetIndex;
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
	 * @param args
	 */
	public static void main(String[] args) {

		MainWindow main = MainWindow.getInstance(); 
		
	    main.buildGUIWindow(main.getShell());
	    
  		main.getShell().open();
  		
  	    while (!main.getShell().isDisposed ()) {
  	        if (!main.getDisplay().readAndDispatch ()) main.getDisplay().sleep();
  	    }
  	    
	    main.getDisplay().dispose();
	}		
}
