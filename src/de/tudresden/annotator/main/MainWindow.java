/**
 * 
 */
package de.tudresden.annotator.main;


import java.io.File;

import org.eclipse.swt.SWT;
import org.eclipse.swt.SWTException;
import org.eclipse.swt.custom.SashForm;
import org.eclipse.swt.graphics.Color;
import org.eclipse.swt.layout.FillLayout;
import org.eclipse.swt.ole.win32.OLE;
import org.eclipse.swt.ole.win32.OleAutomation;
import org.eclipse.swt.ole.win32.OleControlSite;
import org.eclipse.swt.ole.win32.OleFrame;
import org.eclipse.swt.ole.win32.OleListener;
import org.eclipse.swt.widgets.Composite;
import org.eclipse.swt.widgets.Display;
import org.eclipse.swt.widgets.FileDialog;
import org.eclipse.swt.widgets.Label;
import org.eclipse.swt.widgets.MessageBox;
import org.eclipse.swt.widgets.Shell;

import de.tudresden.annotator.oleutils.ApplicationUtils;
import de.tudresden.annotator.oleutils.CommandBarUtils;
import de.tudresden.annotator.oleutils.WorkbookUtils;
import de.tudresden.annotator.oleutils.WorksheetUtils;

/**
 * 
 * @author Elvis Koci
 */
public class MainWindow {
	
	// GUID Event Sink
	private final String IID_AppEvents = "{00024413-0000-0000-C000-000000000046}";
	// Event IDs
	private final int SheetSelectionChange = 0x00000616;
	private final int SheetActivate        = 0x00000619;
	private final int WindowActivate       = 0x00000614;
	private final int WindowDeactivate     = 0x00000615;
	private final int WorkbookActivate     = 0x00000620;
	private final int WorkbookDeactivate   = 0x00000621;
		
	private final Display display = new Display();
	private final Shell shell = new Shell(display);
	
	private OleFrame oleFrame;
	private OleControlSite controlSite;
	
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
	private void buildGUIWindow(){
			
		// this.display.addFilter(SWT.KeyDown, GUIListeners.createArrowButtonPressedEventListener());        		
		// this.display.addFilter(SWT.MouseVerticalWheel, GUIListeners.createMouseWheelEventListener());
		
		// shell properties
		this.shell.setText("Annotator");
	    this.shell.setLayout(new FillLayout());
	    this.shell.setSize(1400, 550);
	    
	    // add listener for the close event ( user clicks the close button (X) )
	    this.shell.addListener(SWT.Close, GUIListeners.createCloseApplicationEventListener());
	    
		// the colors to use for the gui
		Color white = new Color (Display.getCurrent(), 255, 255, 255);
		Color lightGreyShade = new Color (Display.getCurrent(), 247, 247, 247);
		
	    // split shell in two horizontal panels 
	    SashForm mainSash = new SashForm(shell, SWT.HORIZONTAL);
		mainSash.setLayout(new FillLayout());
	
		// the left panel will contain the folder explorer. That is a tree structure of files and folders.
		Composite leftPanel = new Composite(mainSash, SWT.BORDER );
		leftPanel.setLayout(new FillLayout());
		leftPanel.setVisible(true);
		
		Label leftPanelLabel = new Label(leftPanel, SWT.NONE);
		leftPanelLabel.setText("Folder Explorer");
		leftPanelLabel.setBackground(white);
			
		// the right panel will be subdivided into two more panels
		Composite rightPanel = new Composite(mainSash, SWT.BORDER);
		rightPanel.setLayout(new FillLayout());
		
		mainSash.setWeights(new int[] {10,90});
			
		// sub split the right panel 
	    SashForm rightSash = new SashForm(rightPanel, SWT.VERTICAL);
		rightSash.setLayout(new FillLayout());
		
		// create the panel that will embed the excel application
		Composite excelPanel =  new Composite(rightSash, SWT.BORDER );
		excelPanel.setLayout(new FillLayout());
		
		setOleFrame(new OleFrame(excelPanel, SWT.NONE));
		getOleFrame().setBackground(lightGreyShade);
		
		// create the panel that will display the applied annotations for the current file
		Composite annotationsPanel =  new Composite(rightSash, SWT.BORDER );
		annotationsPanel.setLayout(new FillLayout());
		annotationsPanel.setVisible(true);
		
		Label bottomPanelLabel = new Label(annotationsPanel, SWT.NONE);
		bottomPanelLabel.setText("Annotations Panel");
		bottomPanelLabel.setBackground(white);
		
		rightSash.setWeights(new int[] {80,20});
		
		// add a bar menu
	    BarMenu  oleFrameMenuBar = new BarMenu(getOleFrame().getShell());
	    getOleFrame().setFileMenus(oleFrameMenuBar.getMenuItems());		    
	}
	
	protected void setUpWorkbookDisplay( File excelFile){
		
		try {
			setControlSite(new OleControlSite(getOleFrame(), SWT.NONE, "Excel.Sheet", excelFile));
			
			// activate and display excel workbook
			getControlSite().doVerb(OLE.OLEIVERB_INPLACEACTIVATE);
	        
		} catch (IllegalArgumentException iaEx) {
			
			int style = SWT.ICON_ERROR;
			MessageBox message = MainWindow.getInstance().createMessageBox(style);
			message.setMessage("ERROR: Control site could not be created. Illegal argument exception was thrown.");
			message.open();
			MainWindow.getInstance().disposeShell();
			
			// iaEx.printStackTrace();
			System.exit(1);
			
		} catch (SWTException swtEx) {
		
			int style = SWT.ICON_ERROR;
			MessageBox message = MainWindow.getInstance().createMessageBox(style);
			message.setMessage("ERROR: could not embedd the excel workbook. SWT Exception was thrown");
			message.open();
			MainWindow.getInstance().disposeShell();
			
			// swtEx.printStackTrace();
			System.exit(1);
			
		} catch (Exception ex) {
			
			int style = SWT.ICON_ERROR;
			MessageBox message = MainWindow.getInstance().createMessageBox(style);
			message.setMessage("Something went wrong!!! Ensure that you have a version of Microsoft Office Excel"
					+ " installed in your machine. Also, check that the file is not corrupted or wrong format.");
			message.open();
			MainWindow.getInstance().disposeShell();
			
			// ex.printStackTrace();
			System.exit(1);
		}
			       
		setDirectoryPath(excelFile.getParent());
		setFileName(excelFile.getName());
		
		// get excel application as OLE automation object
	    OleAutomation application = ApplicationUtils.getApplicationAutomation(getControlSite());
	    // TODO: suppress alerts
	        
	    // add event listeners
	    OleListener sheetSelectionListener = GUIListeners.createSheetSelectionEventListener();
        getControlSite().addEventListener(application, IID_AppEvents, SheetSelectionChange, sheetSelectionListener);
        
        OleListener sheetActivationlistener = GUIListeners.createSheetActivationEventListener();
        getControlSite().addEventListener(application, IID_AppEvents, SheetActivate, sheetActivationlistener);
                
		// minimize ribbon.	TODO: Try hiding individual CommandBars
	    ApplicationUtils.hideRibbon(application);	
	    
	    // hide menu on right click of user at a worksheet tab
	    CommandBarUtils.setEnabledForCommandBar(application, "Ply", false);
	    
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
	    worksheet.dispose();
	}
				
	/**
	 * @return the display
	 */
	private Display getDisplay() {
		return display;
	}

	/**
	 * @return the shell
	 */
	private Shell getShell() {
		return shell;
	}
	
	/**
	 * Dispose shell
	 */
	protected void disposeShell() {
		if (this.shell != null){
			this.shell.dispose();
		}
	}
	
	/**
	 * give the keyboard focus to the shell
	 */
	protected void setFocusToShell(){
		this.shell.setFocus();
	}
	
	/**
	 * force the keyboard focus to the shell
	 */
	protected void forceFocusToShell(){
		this.shell.forceFocus();
	}
	
	/**
	 * Get OleFrame
	 * @return
	 */
	private OleFrame getOleFrame() {
		return oleFrame;
	}		
	
	/**
	 * Set OleFrame
	 * 
	 * @param oleFrame
	 */
	private void setOleFrame(OleFrame oleFrame) {
		this.oleFrame = oleFrame;
	}
	
	/**
	 * Get OleControlSite
	 * @return
	 */
	private OleControlSite getControlSite() {
		return controlSite;
	}
	
	/**
	 * Set OleControlSite
	 * @param controlSite
	 */
	private void setControlSite(OleControlSite controlSite) {
		this.controlSite = controlSite;
	}
	
	/**
	 * give the keyboard focus to the controlsite
	 */
	protected void setFocusToControlSite(){
		if(controlSite!=null)
			this.controlSite.setFocus();
	}
	
	/**
	 * force the keyboard focus to the controlsite
	 */
	protected void forceFocusToControlSite(){
		if(controlSite!=null)
			this.controlSite.forceFocus();
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
	 * Check if control site is null
	 * @return true if it is null, false otherwise
	 */
	protected boolean isControlSiteNull(){
		return controlSite == null;
	}
	
	/**
	 * Check if control site is dirty
	 * @return true if it is dirty, false otherwise
	 */
	protected boolean isControlSiteDirty(){
		return controlSite.isDirty();
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
	private void setEmbeddedWorkbook(OleAutomation embeddedWorkbook) {
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
	 * Create message box using the "main" window (this class) shell 
	 * @param style one of the relevant SWT constants or their combination
	 * @return a MessageBox object
	 */
	public MessageBox createMessageBox(int style){
		return new MessageBox(this.shell, style);
	}
	
	/**
	 * Create a file dialog using the main shell
	 * @param style one of the relevant SWT constants or their combination
	 * @return FileDialog object
	 */
	public FileDialog createFileDialog(int style){
		return  new FileDialog(this.shell, SWT.OPEN);
	}
	
	/**
	 * @param args
	 */
	public static void main(String[] args) {

		MainWindow main = MainWindow.getInstance(); 
	
	    main.buildGUIWindow();
	    
  		main.getShell().open();
  			 
  	    while (!main.getShell().isDisposed ()) {
  	        if (!main.getDisplay().readAndDispatch ()) main.getDisplay().sleep();
  	    }
  	    
	    main.getDisplay().dispose();
	}		
}
