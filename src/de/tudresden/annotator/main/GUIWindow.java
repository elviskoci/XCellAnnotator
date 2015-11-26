/**
 * 
 */
package de.tudresden.annotator.main;

import java.awt.Container;
import java.io.File;
import java.util.Arrays;

import javax.swing.JComponent;
import javax.swing.LayoutStyle;
import javax.swing.LayoutStyle.ComponentPlacement;

import org.eclipse.swt.SWT;
import org.eclipse.swt.SWTError;
import org.eclipse.swt.events.SelectionAdapter;
import org.eclipse.swt.events.SelectionEvent;
import org.eclipse.swt.layout.FillLayout;
import org.eclipse.swt.layout.GridData;
import org.eclipse.swt.layout.GridLayout;
import org.eclipse.swt.layout.RowLayout;
import org.eclipse.swt.ole.win32.OLE;
import org.eclipse.swt.ole.win32.OleAutomation;
import org.eclipse.swt.ole.win32.OleControlSite;
import org.eclipse.swt.ole.win32.OleEvent;
import org.eclipse.swt.ole.win32.OleFrame;
import org.eclipse.swt.ole.win32.OleListener;
import org.eclipse.swt.ole.win32.Variant;
import org.eclipse.swt.widgets.Button;
import org.eclipse.swt.widgets.Composite;
import org.eclipse.swt.widgets.Control;
import org.eclipse.swt.widgets.Display;
import org.eclipse.swt.widgets.FileDialog;
import org.eclipse.swt.widgets.Group;
import org.eclipse.swt.widgets.Label;
import org.eclipse.swt.widgets.MessageBox;
import org.eclipse.swt.widgets.Shell;

/**
 * 
 * @author Elvis Koci <elvis.koci@tu-dresden.de>
 *
 */
public class GUIWindow {
	
	static final String IID_AppEvents = "{00024413-0000-0000-C000-000000000046}";
	// Event ID
	static final int SheetSelectionChange   = 0x00000616;
	
	OleFrame oleFrame;
	OleControlSite controlSite;
	
	String currentSelection[];
	String activeWorksheetName;
	long activeWorksheetIndex;
	
	/**
	 * @param args
	 */
	public static void main(String[] args) {
	    Display display = new Display ();

	    Shell shell = new Shell (display);
	    
	    GUIWindow gui = new GUIWindow(); 
	    gui.buildGUIWindow(shell);

  		shell.open();
  		
  	    while (!shell.isDisposed ()) {
  	        if (!display.readAndDispatch ()) display.sleep();
  	    }
  	    
	    display.dispose();
	}
	
	
	public void buildGUIWindow(Shell shell){
		
		shell.setText("Annotator");
	    shell.setLayout(new FillLayout());
	    shell.setSize(1200, 600);
	    //shell.setMaximized(true);
	    
		Composite parent = new Composite(shell, SWT.NONE);
		parent.setLayout(new GridLayout(10, true));
	    
		Composite buttons = new Composite(parent, SWT.NONE);
		buttons.setLayout(new GridLayout());
		GridData gridData = new GridData(SWT.BEGINNING, SWT.FILL, false, false);
		buttons.setLayoutData(gridData);
		
		Composite displayArea = new Composite(parent, SWT.BORDER);
		displayArea.setLayout(new FillLayout());
		displayArea.setLayoutData(new GridData(SWT.FILL, SWT.FILL, true, true, 9, 1));
		
		//open new file
		new Label(buttons, SWT.NONE);
		Button openButton = new Button(buttons, SWT.PUSH);
		openButton.setText("Browse...");
		openButton.addSelectionListener(new SelectionAdapter() {
			public void widgetSelected(SelectionEvent e) {
					fileOpen();
					adjustSpreadsheetDisplay();
			}
		});
		
		// Create the radio group that holds the annotation options
		new Label(buttons, SWT.NONE);
	    Group radioButtonGroup = new Group(buttons, SWT.SHADOW_IN);
	    radioButtonGroup.setText("Annotate");
	    radioButtonGroup.setLayout(new RowLayout(SWT.VERTICAL));	    
	    Button metadata = new Button(radioButtonGroup, SWT.RADIO);
	    Button header = new Button(radioButtonGroup, SWT.RADIO);
	    Button data = new Button(radioButtonGroup, SWT.RADIO);
	    Button attributes = new Button(radioButtonGroup, SWT.RADIO);
	    metadata.setText("Metadata");
	    header.setText("Header");
	    data.setText("Data");
	    attributes.setText("Attributes");
	    
	    //  Add annotation button
 		new Label(radioButtonGroup, SWT.NONE);
 		Button addAnnotation = new Button(radioButtonGroup, SWT.PUSH);
 		addAnnotation.setText("Add");
 		addAnnotation.addSelectionListener(new SelectionAdapter() {
 			public void widgetSelected(SelectionEvent e) {
// 				MessageBox msgbox = new MessageBox(shell,SWT.ICON_INFORMATION);
// 				if(currentSelection==null){ 					
// 				     for (Control children : radioButtonGroup.getChildren()) {				    	 
//		 				System.out.println(children);
//		 				    	
//		 				if(children.getStyle()==SWT.RADIO){
//		 				  	System.out.println(((Button)children).getText());
//		 				}
//					 }
// 				     
// 					 msgbox.setMessage("You have annotated the following area/s: {"+currentSelection+"}  as ...");
// 				}else{
// 					 msgbox.setMessage("No areas are currently selected");
// 				}
// 				msgbox.open();
 			}
 		});	
	    	    
		
		oleFrame = new OleFrame(displayArea, SWT.NONE);
	}
	
	/**
	 * Open an excel file for annotation
	 */
	void fileOpen(){
		
		// Select the excel file
		Shell shell = oleFrame.getShell();
		FileDialog dialog = new FileDialog(shell, SWT.OPEN);
		String fileName = dialog.open();
		
		// if no file was selected, return
		if (fileName == null) return;
		
		// dispose OleControlSite if it is not null. 
		disposeControlSite();
				
	    if (controlSite == null) {
			int index = fileName.lastIndexOf('.');
			if (index != -1) {
				String fileExtension = fileName.substring(index + 1);
				if (fileExtension.equalsIgnoreCase("xls") || fileExtension.equalsIgnoreCase("xlsx")) {	
					
					// create a new control site to open the file with Excel
					try {		    	
				        File excelFile = new File(fileName);
				        controlSite = new OleControlSite(oleFrame, SWT.NONE, excelFile); 
				        //controlSite.doVerb(OLE.OLEIVERB_INPLACEACTIVATE);				    
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
	
	
	private void adjustSpreadsheetDisplay(){
		
		if(controlSite==null)
			return;
			
		//get application automation
	    OleAutomation application = getApplicationAutomation(controlSite);
		
	    //Create custom Cell commandbar	
	    createCustomCellCommandBar(application);
	    controlSite.doVerb(OLE.OLEIVERB_INPLACEACTIVATE);	
	    
		//add sheet selection event listener 
	    OleListener listener = createSheetSelectionEventListener(application);
	    controlSite.addEventListener(application, IID_AppEvents, SheetSelectionChange, listener);
		
		//minimize ribbon	
		//TODO: Individual CommandBars
	    hideRibbon(application);	
	    
	    //disable menu on right click of user at a worksheet tab
	    disableTabsCommandBar(application);
	   	    
	    //protect the structure of the active workbook
	    if(!protectActiveWorkbook(application))
	    	System.out.println("\nERROR: Unable to protect active workbook!");
	    
	    //protect all individual worksheets
	    if(!protectAllWorksheets(application))
	    	System.out.println("\nERROR: Unable to protect the worksheets that are part of the active workbook!");
	    
//	    testMacro(application);
	    //TODO: DisplayDocumentInformationPanel (Application)
	    
	    //TODO: DisplayFormulas, Vertical and Horizontal scroll bars, height, width (Window)
	   
//	    application.dispose();
	}
	
	
	/**
	 * Get Excel application automation, OLE object  
	 * @param controlSite
	 * @return
	 */
	private OleAutomation getApplicationAutomation(OleControlSite controlSite){
		
	    OleAutomation excelClient = new OleAutomation(controlSite);
		int[] dispIDs = excelClient.getIDsOfNames(new String[] {"Application"});
		
		if(dispIDs==null){	
			System.out.println("\"Application\" object not found!");
			return null;
		}
		
		Variant pVarResult = excelClient.getProperty(dispIDs[0]);
		if(pVarResult==null){	
			System.out.println("\"Application\" object is null!");
			return null;
		}
		
		OleAutomation application = pVarResult.getAutomation();
		
		pVarResult.dispose();
		excelClient.dispose();
		
		return application;
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
				System.out.print("The selection has changed to: {"+addressVariant.getString()+"}. ");
				currentSelection =  addressVariant.getString().split(",");
				addressVariant.dispose();
				
				int[] areasIds = rangeAutomation.getIDsOfNames(new String[]{"Areas"}); 
				Variant areasVariant = rangeAutomation.getProperty(areasIds[0]);								
				OleAutomation areasAutomation = areasVariant.getAutomation();
				areasVariant.dispose();
				
				int[] countId = areasAutomation.getIDsOfNames(new String[]{"Count"});									
				Variant  countVariant = areasAutomation.getProperty(countId[0]);
				System.out.println("It includes "+countVariant.getString()+" area/s.");
				countVariant.dispose();
				
				args[0].dispose();
				rangeAutomation.dispose();
							
				/*
				 * the second argument is a Worksheet object. get the name and index of the worksheet 	
				 */
				OleAutomation worksheetAutomation = args[1].getAutomation();
				
				int[] nameIds = worksheetAutomation.getIDsOfNames(new String[]{"Name"}); 
				Variant nameVariant = worksheetAutomation.getProperty(nameIds[0]);	
				System.out.print("Selection has occured at worksheet \""+nameVariant.getString()+"\", ");
				activeWorksheetName=nameVariant.getString();
				nameVariant.dispose();
				
				int[] indexIds = worksheetAutomation.getIDsOfNames(new String[]{"Index"}); 
				Variant indexVariant = worksheetAutomation.getProperty(indexIds[0]);	
				System.out.println("which has indexNo "+indexVariant.getString()+".\n");
				activeWorksheetIndex=indexVariant.getLong();
				indexVariant.dispose();
				
				args[1].dispose();
				worksheetAutomation.dispose();
	        }
	    };	       
	    return listener;
	}
	
	
	private boolean testMacro(OleAutomation application) {
		
		int[] ee4mIds = application.getIDsOfNames(new String[]{"ExecuteExcel4Macro"});
				
		Variant[] parameters = new Variant[1];
	    parameters[0] = new Variant("Dim ContextMenu As CommandBar\n"+
		   "Dim MySubMenu As CommandBarControl\n"+
		   "Call DeleteFromCellMenu\n"+
		   "Set ContextMenu = Application.CommandBars(\"Cell\")\n"+
		   "ContextMenu.Controls.Add Type:=msoControlButton, ID:=3, before:=1");
	    
	    Variant result = application.invoke(ee4mIds[0],parameters);
	    parameters[0].dispose();
	    
	    boolean isSuccess = false;
	    if(result!=null){
	    	isSuccess = true;
	    	result.dispose();	
	    }
	   
	    System.out.println("Test Macro "+isSuccess+" \n");
	    
	    return isSuccess;
	}
	
	
	/**
	 * Hide Ribbon from Excel UI
	 * @param application
	 * @return
	 */
	private boolean hideRibbon(OleAutomation application){
		
		int[] ee4mIds = application.getIDsOfNames(new String[]{"ExecuteExcel4Macro"});
		
		Variant[] parameters = new Variant[1];
	    parameters[0] = new Variant("SHOW.TOOLBAR(\"Ribbon\",False)");
	    
	    Variant result = application.invoke(ee4mIds[0],parameters);
	    //System.out.println("\nThe result of ExecuteExcel4Macro method invocation: "+result);
	   
	    boolean isSuccess = false;
	    if(result!=null)
	    	isSuccess = true;
	    
	    parameters[0].dispose();
	    result.dispose();	
	    
	    return isSuccess;
	}
	
	
	private void createCustomCellCommandBar(OleAutomation application){
		
		OleAutomation cellCBAutomation = getCellCommandBar(application);
				
		if(cellCBAutomation==null)
			return;
		
		int[] controlsPropertyIds = cellCBAutomation.getIDsOfNames(new String[]{"Controls"});
		Variant controlsVariant = cellCBAutomation.getProperty(controlsPropertyIds[0]);
		OleAutomation contolsAutomation = controlsVariant.getAutomation();
		controlsVariant.dispose();
		
		//make existing controls not visible
		int[] itemPropertyIds = contolsAutomation.getIDsOfNames(new String[]{"Item"});
	
		Variant[] parameters = new Variant[1];
		parameters[0] = new Variant(1);
		Variant controlItemVariant = contolsAutomation.getProperty(itemPropertyIds[0],parameters);
		parameters[0].dispose();
		
		int i=1;
		while (controlItemVariant!=null) {			
			OleAutomation controlItemAutomation = controlItemVariant.getAutomation();
			int[] visiblePropertyIds = controlItemAutomation.getIDsOfNames(new String[]{"Visible"});
			controlItemAutomation.setProperty(visiblePropertyIds[0],new Variant(false));
			parameters[0] = new Variant(i++);
			controlItemVariant.dispose();
			controlItemVariant = contolsAutomation.getProperty(itemPropertyIds[0],parameters);
			parameters[0].dispose();
		}
		
		// create custom control item
		int[] addMethodIds = contolsAutomation.getIDsOfNames(new String[]{"Add"});
		if(addMethodIds==null){
			System.out.println("addMethodIds is null");
			return;
		}else{
			System.out.println(Arrays.toString(addMethodIds));
		}
		
		Variant[] args = new Variant[1];
		args[0] = new Variant();
		//args[0] = new Variant("msoControlPopup");
		//args[0] = new Variant(1);
		Variant myControlItemVariant =  contolsAutomation.invoke(addMethodIds[0]);
		if(myControlItemVariant==null){
			System.out.println("myControlItemVariant is null");
			return;
		}
		OleAutomation myControlItem = myControlItemVariant.getAutomation();
		myControlItemVariant.dispose();
		
		int[] captionProperyIds = myControlItem.getIDsOfNames(new String[]{"Caption"});
		int[] tagProperyIds = myControlItem.getIDsOfNames(new String[]{"Tag"});
		myControlItem.setProperty(captionProperyIds[0], new Variant("Annotate"));
		myControlItem.setProperty(tagProperyIds[0], new Variant("annotation_controls"));
	}
	
	
	/**
	 * Get the Cell command bar automation
 	 * @param application
	 * @return
	 */
	private OleAutomation getCellCommandBar(OleAutomation application) {
		
		int[] commandBarsPropertyIds = application.getIDsOfNames(new String[]{"CommandBars"});
		if (commandBarsPropertyIds == null) {
			System.out.println("Property \"CommandBars\" of \"Application\" OLE Object is null!");
			return null;
		}
		
		Variant commandBarsVariant =  application.getProperty(commandBarsPropertyIds[0]);	
		if(commandBarsVariant == null){
			System.out.println("\"CommandBars\" variant is null!");
			return null;		
		}
		OleAutomation commandBarsAutomation = commandBarsVariant.getAutomation();
		commandBarsVariant.dispose();
			
		int[] itemPropertyIds = commandBarsAutomation.getIDsOfNames(new String[]{"Item"});
		if(itemPropertyIds == null){
			System.out.println("Property \"Item\" of \"CommandBars\" OLE object not found!");
			return null;
		}

		Variant[] parameters = new Variant[1];
		parameters[0] = new Variant("cell");
		Variant cellCommandBar = commandBarsAutomation.getProperty(itemPropertyIds[0],parameters);
		parameters[0].dispose();
		
		if(cellCommandBar==null){
			System.out.println("There is no CommandBar named \"cell\"");
			return null;
		}
		OleAutomation cellCBAutomation = cellCommandBar.getAutomation();
		cellCommandBar.dispose();
		
		int[] addMethodIds = commandBarsAutomation.getIDsOfNames(new String[]{"Add"});
		if(addMethodIds==null){
			System.out.println("addMethodIds is null");
		}else{
			System.out.println(Arrays.toString(addMethodIds));
		}
		
		Variant[] args = new Variant[1];
		args[0] = new Variant();
		Variant myControlItemVariant =  commandBarsAutomation.invoke(addMethodIds[0]);
		if(myControlItemVariant==null){
			System.out.println("myControlItemVariant is null");
		}else{
			OleAutomation myControlItem = myControlItemVariant.getAutomation();
			myControlItemVariant.dispose();
		}
		
		return cellCBAutomation;
	}
	
	/**
	 * Disable the menu that is displayed when right click on workbook tabs 
 	 * @param application
	 * @return
	 */
	private boolean disableTabsCommandBar(OleAutomation application) {
		
		int[] commandBarsPropertyIds = application.getIDsOfNames(new String[]{"CommandBars"});
		if (commandBarsPropertyIds == null) {
			System.out.println("Property \"CommandBars\" of \"Application\" OLE Object is null!");
			return false;
		}
		
		Variant commandBarsVariant =  application.getProperty(commandBarsPropertyIds[0]);	
		if(commandBarsVariant == null){
			System.out.println("\"CommandBars\" variant is null!");
			return false;		
		}
		OleAutomation commandBarsAutomation = commandBarsVariant.getAutomation();
		commandBarsVariant.dispose();
			
		int[] itemPropertyIds = commandBarsAutomation.getIDsOfNames(new String[]{"Item"});
		if(itemPropertyIds == null){
			System.out.println("Property \"Item\" of \"CommandBars\" OLE object not found!");
			return false;
		}

		Variant[] parameters = new Variant[1];
		parameters[0] = new Variant("Ply");
		Variant tabsCBVariant = commandBarsAutomation.getProperty(itemPropertyIds[0],parameters);
		parameters[0].dispose();
		
		if(tabsCBVariant==null){
			System.out.println("There is no CommandBar named \"Workbook tabs\"");
			return false;
		}
		OleAutomation tabsCBAutomation = tabsCBVariant.getAutomation();
		tabsCBVariant.dispose();
		
		int[] enabledPropertyIds = tabsCBAutomation.getIDsOfNames(new String[]{"Enabled"});
		if(enabledPropertyIds == null){
			System.out.println("Property \"Enabled\" of \"CommandBars\" OLE object not found!");
			return false;
		}
		
		boolean isSuccess = tabsCBAutomation.setProperty(enabledPropertyIds[0], new Variant(false));
		return isSuccess;
	}
	
	
	/**
	 * get ole automation for the active workbook 
	 * 
	 * @param application
	 * @return
	 */
	private OleAutomation getActiveWorkbook(OleAutomation application){
		
		int[] workbookIds = application.getIDsOfNames(new String[]{"ActiveWorkbook"});	
		if (workbookIds == null){			
			System.out.println("\"ActiveWorkbook\" property not found for \"Application\" object!");
			return null;
		}		
		Variant workbookVariant = application.getProperty(workbookIds[0]);
		if (workbookVariant == null) {
			System.out.println("Workbook variant is null!");
			return null;
		}		
		OleAutomation workbookAutomation =  workbookVariant.getAutomation();
		workbookVariant.dispose();
		
		return workbookAutomation;
	}
	
	
	/**
	 * Protect the structure of the active workbook 
	 * @param application
	 * @return
	 */
	private boolean protectActiveWorkbook(OleAutomation application){
		
		// get ole automation for the active workbook 
		OleAutomation workbookAutomation = getActiveWorkbook(application);
		if (workbookAutomation==null)
			return false;
		
		// invoke the "Protect" method for the active workbook
		int[] protectMethodIds = workbookAutomation.getIDsOfNames(new String[]{"Protect", "Structure", "Windows"});
		if (protectMethodIds == null) {
			System.out.println("Method \"Protect\" not found for \"Workbook\" object!");
			return false;
		}else{
			Variant[] args = new Variant[2];
			args[0] = new Variant(true);
			args[1] = new Variant(false);
			
			Variant result = workbookAutomation.invoke(protectMethodIds[0],args,Arrays.copyOfRange(protectMethodIds, 1, protectMethodIds.length));	
			//System.out.println("Result of Workbook.Protect(): "+result);
			if(result==null)
				return false;
			
			result.dispose();
			for (Variant arg: args) {
				arg.dispose();
			}
		}
		
		workbookAutomation.dispose();
		return true;
	}
	
	
	/**
	 * Unprotect the structure of the active workbook
	 * @param application
	 * @return
	 */
	private boolean unprotectActiveWorkbook(OleAutomation application){
		
		// get ole automation for the active workbook 
		OleAutomation workbookAutomation = getActiveWorkbook(application);
		if (workbookAutomation==null)
			return false;
				
		// invoke the "Protect" method for the active workbook
		int[] unprotectMethodIds = workbookAutomation.getIDsOfNames(new String[]{"Unprotect"});
		if (unprotectMethodIds == null) {
			System.out.println("Method \"Unprotect\" not found for \"Workbook\" object!");
			return false;
		}else{
			Variant[] args = new Variant[1];
			args[0] = new Variant();
			
			Variant result = workbookAutomation.invoke(unprotectMethodIds[0],args);	
			//System.out.println("Result of Workbook.Unprotect(): "+result);
			if(result==null)
				return false;
			
			result.dispose();
			args[0].dispose();
		}
		
		workbookAutomation.dispose();
		return true;
	}
	
	
	/**
	 * Protect all worksheet that are in the active worksheet
	 * @param application
	 * @return
	 */
	private boolean protectAllWorksheets(OleAutomation application){
		
		// mark each worksheet as protected 
		int[] worksheetsObjectIds = application.getIDsOfNames(new String[]{"Worksheets"});
		if (worksheetsObjectIds == null) {
			System.out.println("Property \"Worksheets\" of \"Application\" OLE Object is null!");
			return false;
		}
		
		Variant worksheetsVariant =  application.getProperty(worksheetsObjectIds[0]);	
		if(worksheetsVariant == null){
			System.out.println("\"Worksheets\" variant is null!");
			return false;		
		}
		OleAutomation worksheetsAutomation = worksheetsVariant.getAutomation();
		worksheetsVariant.dispose();
		
		int[] countProperyIds = worksheetsAutomation.getIDsOfNames(new String[]{"Count"});
		if(countProperyIds == null){
			System.out.println("Property \"Count\" of \"Worksheets\" OLE object is null!");
			return false;
		}
				
		Variant countPropertyVariant =  worksheetsAutomation.getProperty(countProperyIds[0]);
		if(countPropertyVariant == null){
			System.out.println("\"Count\" variant is null!");
			return false;
		}				
				
		int count = countPropertyVariant.getInt();
		countPropertyVariant.dispose();
		
		int[] itemPropertyIds = worksheetsAutomation.getIDsOfNames(new String[]{"Item"});
		if(itemPropertyIds == null){
			System.out.println("Property \"Item\" of \"Worksheets\" OLE object not found!");
			return false;
		}

		for (int i = 1; i <= count; i++) {
			Variant[] args = new Variant[1];
			args[0] = new Variant(i);		
			Variant nextWorkbookVariant = worksheetsAutomation.getProperty(itemPropertyIds[0],args);	
			if(!protectWorksheet(nextWorkbookVariant.getAutomation())){
				System.out.println("ERROR: Could not protect one of the workbooks!");
				return false;
			}
			nextWorkbookVariant.dispose();
			args[0].dispose();
		}
		
		return true;
	}
	
	
	/**
	 * Protect the data, formating, and structure of the specified worksheet
	 * @param worksheetAutomation
	 * @return
	 */
	private boolean protectWorksheet(OleAutomation worksheetAutomation){
		
		// get the id of the "Protect" method and the considered parameters
		// you can find the documentation of this OLE method here https://msdn.microsoft.com/EN-US/library/ff840611.aspx
		int[] protectMethodIds = worksheetAutomation.getIDsOfNames(new String[]{"Protect", "DrawingObjects", "Contents", "Scenarios", 
				"AllowFormattingCells", "AllowFormattingColumns", "AllowFormattingRows", "AllowInsertingColumns", "AllowInsertingRows", 
				"AllowInsertingHyperlinks", "AllowDeletingColumns", "AllowDeletingRows", "AllowSorting", "AllowFiltering", "AllowUsingPivotTables" });
		
		if (protectMethodIds == null) {
			System.out.println("Method \"Protect\" of \"Worksheet\" OLE Object is not found!");
			return false;
		}else{
			Variant[] args = new Variant[14];
			args[0] = new Variant(true);
			args[1] = new Variant(true);
			args[2] = new Variant(true);
			args[3] = new Variant(false);
			args[4] = new Variant(false);
			args[5] = new Variant(false);
			args[6] = new Variant(false);
			args[7] = new Variant(false);
			args[8] = new Variant(false);
			args[9] = new Variant(false);
			args[10] = new Variant(false);
			args[11] = new Variant(false);
			args[12] = new Variant(false);
			args[13] = new Variant(false);
			
			Variant result = worksheetAutomation.invoke(protectMethodIds[0],args,Arrays.copyOfRange(protectMethodIds, 1, protectMethodIds.length));	
			result.dispose();
			for (Variant arg: args) {
				arg.dispose();
			}
		}				
		worksheetAutomation.dispose();	
		return true;
	}
	
	
	/**
	 * Dispose control site 
	 */
	void disposeControlSite() {
		if (controlSite != null){
		
			OleAutomation application=  getApplicationAutomation(controlSite);
			if(!unprotectActiveWorkbook(application)){
				System.out.println("\nERROR: Failed to unprotect active workbook!");
			}
			application.dispose();
			controlSite.dispose();
		}
		controlSite = null;
	}
	
}
