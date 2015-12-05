/**
 * 
 */
package de.tudresden.annotator.main;

import java.util.Arrays;

import org.eclipse.swt.ole.win32.OleAutomation;
import org.eclipse.swt.ole.win32.OleControlSite;
import org.eclipse.swt.ole.win32.Variant;

/**
 * @author Elvis
 *
 */
public class ExcelUIModifier {
		
	/**
	 * Protect the structure of the active workbook 
	 * @param application
	 * @return
	 */
	public static boolean protectWorkbook(OleAutomation workbookAutomation){
		
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
//			System.out.println("Result of Workbook.Protect(): "+result);
			if(result==null)
				return false;
			
			result.dispose();
			for (Variant arg: args) {
				arg.dispose();
			}
		}
		
		return true;
	}
	
	
   /**
	 * Unprotect the structure of the active workbook
	 * @param application
	 * @return
	 */
	public static boolean unprotectWorkbook(OleAutomation workbookAutomation){
		
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
//			System.out.println("Result of Workbook.Unprotect(): "+result);
			if(result==null)
				return false;
			
			result.dispose();
			args[0].dispose();
		}
		
		return true;
	}
	

	/**
	 * Protect all worksheet that are part of the given workbook
	 * @param application
	 * @return
	 */
	public static boolean protectAllWorksheets(OleAutomation workbookAutomation){
		
		OleAutomation worksheetsAutomation = AutomationUtils.getWorksheetsAutomation(workbookAutomation);

		int count = AutomationUtils.getNumberOfObjectsInOleCollection(worksheetsAutomation);
		
		int i; 
		boolean isSuccess=true; 
		for (i = 1; i <= count; i++) {
		
			OleAutomation nextWorkbookAutomation = AutomationUtils.getItem(worksheetsAutomation, i);					
			if(!protectWorksheet(nextWorkbookAutomation)){
				System.out.println("ERROR: Could not protect one of the workbooks!");
				isSuccess=false;			
			}	
			nextWorkbookAutomation.dispose();	
			if(!isSuccess){
				break;
			}
		}	
		
		if(!isSuccess){
			for(int j=1; j<i;j++){
				OleAutomation nextWorkbookAutomation =  AutomationUtils.getItem(worksheetsAutomation, j);
				unprotectWorksheet(nextWorkbookAutomation);
				nextWorkbookAutomation.dispose();
			}
			worksheetsAutomation.dispose();
			return false;
		}
		
		worksheetsAutomation.dispose();
		return true;
	}
	
	
	/**
	 * Protect the data, formating, and structure of the specified worksheet
	 * @param worksheetAutomation
	 * @return
	 */
	public static boolean protectWorksheet(OleAutomation worksheetAutomation){
		
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
	 * Unprotect the specified worksheet 
	 * 
	 * @param worksheetAutomation
	 * @return
	 */
	public static boolean unprotectWorksheet(OleAutomation worksheetAutomation){
		
		// get the id of the "Unprotect" method for worksheet OLE object 
		int[] unprotectMethodIds = worksheetAutomation.getIDsOfNames(new String[]{"Unprotect"});
		if(unprotectMethodIds==null){
			System.out.println("Method \"Unprotect\" of \"Worksheet\" OLE Object is not found!");
			return false;
		}
		
		// call the unprotect method  
		Variant result = worksheetAutomation.invoke(unprotectMethodIds[0]);

		if(result==Variant.NULL){
			result.dispose();
			return false;
		}	
		
		result.dispose();
		return true;
	}
	
	/**
	 * Hide Ribbon from Excel GUI
	 * @param application
	 * @return
	 */
	public static boolean hideRibbon(OleAutomation application){
		
		int[] ee4mIds = application.getIDsOfNames(new String[]{"ExecuteExcel4Macro"});
		
		Variant[] parameters = new Variant[1];
	    parameters[0] = new Variant("SHOW.TOOLBAR(\"Ribbon\",False)");
	    
	    Variant result = application.invoke(ee4mIds[0],parameters);
//	    System.out.println("\nThe result of ExecuteExcel4Macro method invocation: "+result);
	   
	    boolean isSuccess = false;
	    if(result!=null)
	    	isSuccess = true;
	    
	    parameters[0].dispose();
	    result.dispose();	
	    
	    return isSuccess;
	}
	
	
	/**
	 * Create a custom "Cell" command bar. All the exiting controls (menu items) will be hidden, and new ones will be added.
	 * @param application
	 */
	public static void createCustomCellCommandBar(OleAutomation application){
		
		// get the "Cell" command bar automation
		OleAutomation cellCBAutomation = CommandBarsHelper.getCommandBarByName(application,"cell");		
		if(cellCBAutomation==null)
			return;
		
		// get CommandBarsControls object automation.
		OleAutomation contolsAutomation = CommandBarsHelper.getCommandBarControls(cellCBAutomation);
		cellCBAutomation.dispose();
		
		// temporary delete the (menu) items in the "Cell" command bar. This controls will appear again when Excel application is started  
		CommandBarsHelper.deleteControlsTemporary(contolsAutomation);
		
		// add new CommandBarPopup control
		int[] addMethodIds = contolsAutomation.getIDsOfNames(new String[]{"Add", "Type", "Before"});
		
		Variant[] args = new Variant[addMethodIds.length-1];
		args[0] = new Variant(10);
		args[1] = new Variant(1);
		
		Variant myControlItemVariant = contolsAutomation.invoke(addMethodIds[0],args,Arrays.copyOfRange(addMethodIds, 1, addMethodIds.length));
		OleAutomation myPopUpControl = myControlItemVariant.getAutomation();
		myControlItemVariant.dispose();	
		for (Variant arg : args) {
			arg.dispose();
		}
		
		// set the properties of the control
		int[] captionProperyIds = myPopUpControl.getIDsOfNames(new String[]{"Caption"});
		int[] tagProperyIds = myPopUpControl.getIDsOfNames(new String[]{"Tag"});
		myPopUpControl.setProperty(captionProperyIds[0], new Variant("Annotate as"));
		myPopUpControl.setProperty(tagProperyIds[0], new Variant("annotation_controls"));
		
		// add sub-controls (sub-menus) 
		OleAutomation mySubContolsAutomation = CommandBarsHelper.getCommandBarControls(myPopUpControl);
		addMethodIds = contolsAutomation.getIDsOfNames(new String[]{"Add", "Type"});
		args = new Variant[addMethodIds.length-1];
		args[0] = new Variant(1);
		
		String[] captions = new String[]{"Table","Metadata","Header","Attributes","Data"};
		Variant[] mySubControlVariant = new Variant[captions.length];
		OleAutomation mySubControl;
		for (int i = 0; i < captions.length; i++) {
			mySubControlVariant[i] = mySubContolsAutomation.invoke(addMethodIds[0], args, Arrays.copyOfRange(addMethodIds, 1, addMethodIds.length));		
			
			mySubControl= mySubControlVariant[i].getAutomation(); 
			int[] captionPropetyIds = mySubControl.getIDsOfNames(new String[]{"Caption"});
			mySubControl.setProperty(captionPropetyIds[0], new Variant(captions[i]));
			
//			int[] onActionPropertyIds = mySubControl.getIDsOfNames(new String[]{"OnAction"});
//			mySubControl.setProperty(onActionPropertyIds[0], new Variant("MsgBox \"You annotated as ...\""));
//			
//			MainWindow.getInstance().getControlSite().addPropertyListener(onActionPropertyIds[0], new OleListener() {			
//				@Override
//				public void handleEvent(OleEvent event) {
//					System.out.println("Event Captured!!!");
//				}
//			});
			
			mySubControlVariant[i].dispose();
			mySubControl.dispose();
		}
		contolsAutomation.dispose();
		mySubContolsAutomation.dispose();
		for (Variant arg : args) {
			arg.dispose();
		}	
	}
	
	
	/**
	 * Undo changes done to the "Cell" commandbar during the current session
	 * @param application
	 */
	public static void undoChangesToCellCommandBar(OleAutomation application){
		
	    // get the "Cell" command bar automation
		OleAutomation cellCBAutomation = CommandBarsHelper.getCommandBarByName(application,"cell");		
		if(cellCBAutomation==null)
			return;
		
		// get CommandBarsControls object automation.
		OleAutomation contolsAutomation = CommandBarsHelper.getCommandBarControls(cellCBAutomation);
		cellCBAutomation.dispose();
		
		// show (Make visible) all the control (menu) items in the "Cell" command bar
		CommandBarsHelper.setVisibilityOfControls(contolsAutomation, true);
		// delete custom controls, created during the current session of the application 
		CommandBarsHelper.deleteCustomControlsByTag(contolsAutomation,"annotation_controls");		
	}
	
	
	/** 
	 * Annotate the selected areas by drawing a border around each one of them 
	 * @param colorIndex
	 */
	public static void annotateByBorderingSelectedAreas(int colorIndex){
		 
		String name = MainWindow.getInstance().getActiveWorksheetName();
		String[] selectedAreas =  MainWindow.getInstance().getCurrentSelection();
		
		// get the OleAutomation object for the active worksheet using its name
		OleAutomation worksheetAutomation = AutomationUtils.getWorksheetAutomationByName(name);
	
		// unprotect the worksheet in order to change the border for the range 
		unprotectWorksheet(worksheetAutomation);
		
		// for each area in the range draw a border
		for (String area : selectedAreas) {
			String[] subStrings = area.split(":");
			
			String topRightCell = subStrings[0];
			String downLeftCell = null; 
			if(subStrings.length==2)
				downLeftCell = subStrings[1];
			
			OleAutomation rangeAutomation = AutomationUtils.getRangeAutomation(worksheetAutomation, topRightCell, downLeftCell);
			
			drawBorderAroundRange(rangeAutomation,1,4,colorIndex);
			rangeAutomation.dispose();
		}
		
		// protect the workbook to prevent the user from modifying the content of the sheet
		protectWorksheet(worksheetAutomation);
		worksheetAutomation.dispose();
	
		
		// unprotect the worksheet in order to change the border for the range 
		// unprotectWorksheet( MainWindow.getInstance().getActiveWorksheetAutomation() );
		
		// set the specified border around the selected areas
		// setBorderToRange( MainWindow.getInstance().getSelectedRangeAutomation(), 1, 4, colorIndex );
		
		// protect the workbook to prevent the user from modifying the content of the sheet
		// protectWorksheet( MainWindow.getInstance().getActiveWorksheetAutomation() );
	
	}
	

	/**
	 * Annotate the selected areas by drawing textbox on top of each one of them.
	 * The color of the textbox depends on the Annotation Class.
	 * 
	 * @param color a long that represents a RGB color. Is calculated as B * 65536 + G * 256 + R
	 */
	public static void annotateSelectedAreasWithTextboxes(long color, String label){
		
		String name = MainWindow.getInstance().getActiveWorksheetName();
		String[] selectedAreas =  MainWindow.getInstance().getCurrentSelection();
		
		// get the OleAutomation object for the active worksheet using its name
		OleAutomation worksheetAutomation = AutomationUtils.getWorksheetAutomationByName(name);
	
		// unprotect the worksheet in order to apply changes
		unprotectWorksheet(worksheetAutomation);
		
		// for each area in the range draw a textbox
		for (String area : selectedAreas) {
			String[] subStrings = area.split(":");
			
			String topRightCell = subStrings[0];
			String downLeftCell = null; 
			if(subStrings.length==2)
				downLeftCell = subStrings[1];
			
			OleAutomation rangeAutomation = AutomationUtils.getRangeAutomation(worksheetAutomation, topRightCell, downLeftCell);
			
			double left = AutomationUtils.getRangeLeftPosition(rangeAutomation);
			double top = AutomationUtils.getRangeTopPosition(rangeAutomation);
			double width = AutomationUtils.getRangeWidth(rangeAutomation);
			double height = AutomationUtils.getRangeHeight(rangeAutomation);
			
			OleAutomation textboxAutomation = drawTextBox(worksheetAutomation, left, top, width, height); 
			OleAutomation fillFormatAutomation = getFillFormatAutomation(textboxAutomation);
			OleAutomation textFrameAutomation = getTextFrameAutomation(textboxAutomation);
			OleAutomation charactersAutomation = getCharactersAutomation(textFrameAutomation);
			OleAutomation fontAutomation = getFontAutomation(charactersAutomation);
			
			setShapeBackgroundColor(fillFormatAutomation, color);
			setShapeFillTransparency(fillFormatAutomation, 0.60);		
			setText(charactersAutomation, label);
			setHorizontalAlignment(textFrameAutomation, -4108); // align center 
			setVerticalAlignment(textFrameAutomation, -4108); // align center
			
			long whiteColor = 255 * 65536 + 255 * 256 + 255;
			setFontColor(fontAutomation, whiteColor);
			setBoldFont(fontAutomation, true); 
			setFontSize(fontAutomation, 11); // TODO: should be relative to the size of the range 
			
			rangeAutomation.dispose();
			fillFormatAutomation.dispose();
			fontAutomation.dispose();
			charactersAutomation.dispose();
			textFrameAutomation.dispose();
			textboxAutomation.dispose();
		}
		
		// protect the workbook to prevent the user from modifying it
		protectWorksheet(worksheetAutomation);
		worksheetAutomation.dispose();
	}
	
	
	/**
	 * Create a border around the range with the specified characteristics
	 * 
	 * @param rangeAutomation
	 * @param lineStyle
	 * @param weight
	 * @param colorIndex
	 */
	public static void  drawBorderAroundRange(OleAutomation rangeAutomation, int lineStyle, int weight, int colorIndex){
		
		//  set border around the selected range 
		int[] borderAroundMethodIds = rangeAutomation.getIDsOfNames(new String[]{"BorderAround","LineStyle", "Weight", "ColorIndex"}); // "Color"
		Variant methodParams[] = new Variant[3];
		methodParams[0] = new Variant(lineStyle); // line style (e.g., continuous, dashed ) 
		methodParams[1] = new Variant(weight); // border weight  (e.g., thick )
		methodParams[2] = new Variant(colorIndex); // Index into the current color
	
		int[] paramIds = Arrays.copyOfRange(borderAroundMethodIds, 1, borderAroundMethodIds.length);
		rangeAutomation.invoke(borderAroundMethodIds[0], methodParams, paramIds);
		
		for (Variant v : methodParams) {
			v.dispose();
		}		
	}
	
	/**
	 * Draw textbox at the specified location in the worksheet 
	 * @param sheetAutomation
	 * @param left
	 * @param top
	 * @param width
	 * @param height
	 * @return
	 */
	public static OleAutomation drawTextBox(OleAutomation sheetAutomation, double left, double top, double width, double height){
		
		OleAutomation shapesAutomation = AutomationUtils.getWorksheetShapes(sheetAutomation);
		
		//  set border around the selected range 
		int[] addTextboxMethodIds = shapesAutomation.getIDsOfNames(new String[]{"AddTextbox", "Orientation", "Left", "Top", "Width", "Height"}); 
		Variant methodParams[] = new Variant[5];
		methodParams[0] = new Variant(1);
		methodParams[1] = new Variant(left+0.5); 
		methodParams[2] = new Variant(top+0.5); 
		methodParams[3] = new Variant(width-1); 
		methodParams[4] = new Variant(height-1);	
		
		Variant textboxVariant = shapesAutomation.invoke(addTextboxMethodIds[0],methodParams);
		
		shapesAutomation.dispose();
		for (Variant v : methodParams) {
			v.dispose();
		}
		
		OleAutomation textboxAutomation = null;
		if(textboxVariant!=null){
			textboxAutomation = textboxVariant.getAutomation();
			textboxVariant.dispose();
		}else{
			System.out.println("ERROR: Failed to draw textbox annotation!!!");
			System.exit(1);
		}
		
		return textboxAutomation;
	}
	

	/**
	 * Get FillFormat OleAutomation. This object can be used to change the format of the shape fill 
	 * @param shapeAutomation
	 * @return
	 */
	public static OleAutomation getFillFormatAutomation(OleAutomation shapeAutomation){
		
		int[] fillPropertyIds = shapeAutomation.getIDsOfNames(new String[]{"Fill"}); 
		Variant fillFormatVariant = shapeAutomation.getProperty(fillPropertyIds[0]);
		OleAutomation fillFormatAutomation = fillFormatVariant.getAutomation();
		fillFormatVariant.dispose();
		
		return fillFormatAutomation;
	}
		
	/**
	 * Set the color of the shape fill 
	 * @param fillFormatAutomation
	 * @param color
	 * @return
	 */
	public static boolean setShapeBackgroundColor(OleAutomation fillFormatAutomation, long color){
	
		int[] foreColorPropertyIds = fillFormatAutomation.getIDsOfNames(new String[]{"ForeColor"}); // BackColor does not work
		Variant foreColorVariant = fillFormatAutomation.getProperty(foreColorPropertyIds[0]);
		OleAutomation foreColorAutomation = foreColorVariant.getAutomation();

		int[] rgbPropertyIds = foreColorAutomation.getIDsOfNames(new String[]{"RGB"}); //alternative "SchemeColor" 
		foreColorAutomation.setProperty(rgbPropertyIds[0], new Variant(color)); 
	
		boolean isSuccess = fillFormatAutomation.setProperty(foreColorPropertyIds[0], foreColorVariant);			
		foreColorVariant.dispose();
		foreColorAutomation.dispose();
		
		return isSuccess;
	}
	
	/**
	 * Set the transparency of the shape fill
	 * @param fillFormatAutomation
	 * @param transparency
	 * @return
	 */
	public static boolean setShapeFillTransparency(OleAutomation fillFormatAutomation, double transparency){
		int[] transparencyPropertyIds = fillFormatAutomation.getIDsOfNames(new String[]{"Transparency"}); 
		return fillFormatAutomation.setProperty(transparencyPropertyIds[0], new Variant(transparency));	
	}
	
	/**
	 * Get the TextFrame OleAutomation. This object can be used to manage text in a shape.
	 * @param shapeAutomation
	 * @return
	 */
	public static OleAutomation getTextFrameAutomation(OleAutomation shapeAutomation){
		
		int[] textFramePropertyIds = shapeAutomation.getIDsOfNames(new String[]{"TextFrame"}); 
		Variant textFrameVariant = shapeAutomation.getProperty(textFramePropertyIds[0]);
		OleAutomation textFrameAutomation = textFrameVariant.getAutomation();
		textFrameVariant.dispose();
		
		return textFrameAutomation;
	}
	
	/**
	 * Set text horizontal alignment in the shape TextFrame  
	 * @param textFrameAutomation
	 * @param alignment
	 * @return
	 */
	public static boolean setVerticalAlignment(OleAutomation textFrameAutomation, int alignment){
		
		int[] verticalAlignmentPropertyIds = textFrameAutomation.getIDsOfNames(new String[]{"VerticalAlignment"});
		return textFrameAutomation.setProperty(verticalAlignmentPropertyIds[0], new Variant(alignment));
	}
	
	/**
	 * Set text vertical alignment in the shape TextFrame
	 * @param textFrameAutomation
	 * @param alignment
	 * @return
	 */
	public static boolean setHorizontalAlignment(OleAutomation textFrameAutomation, int alignment){
		
		int[] horizontalAlignmentPropertyIds = textFrameAutomation.getIDsOfNames(new String[]{"HorizontalAlignment"});
		return textFrameAutomation.setProperty(horizontalAlignmentPropertyIds[0], new Variant(alignment));
	}

	/**
	 * Get Characters OleAutomation. Use this object to set text and handle its format 
	 * @param textFrameAutomation
	 * @return
	 */
	public static OleAutomation getCharactersAutomation(OleAutomation textFrameAutomation){
		
		int[] charactersPropertyIds = textFrameAutomation.getIDsOfNames(new String[]{"Characters"});
		Variant charactersVariant = textFrameAutomation.invoke(charactersPropertyIds[0]);
		OleAutomation charactersAutomation = charactersVariant.getAutomation();
		charactersVariant.dispose();
		
		return charactersAutomation;
	}
	
	/**
	 * Set Text using an Characters OleAutomation
	 * @param charactersAutomation
	 * @param text
	 * @return
	 */
	public static boolean setText(OleAutomation charactersAutomation, String text){
		
		int[] textPropertyIds = charactersAutomation.getIDsOfNames(new String[]{"Text"}); 
		return charactersAutomation.setProperty(textPropertyIds[0], new Variant(text));
	}
	
	/**
	 * Get Font OleObject. This object can be used to change text font attributes
	 * @param charactersAutomation
	 * @return
	 */
	public static OleAutomation getFontAutomation(OleAutomation charactersAutomation){
		
		int[] fontPropertyIds = charactersAutomation.getIDsOfNames(new String[]{"Font"});
		Variant fontVariant = charactersAutomation.getProperty(fontPropertyIds[0]);
		OleAutomation fontAutomation = fontVariant.getAutomation();
		fontVariant.dispose();
		
		return fontAutomation;
	}
	
	/**
	 * Set text font size
	 * @param fontAutomation
	 * @param size
	 * @return
	 */
	public static boolean setFontSize(OleAutomation fontAutomation, int size){		
		int[] sizePropertyIds = fontAutomation.getIDsOfNames(new String[]{"Size"}); 
		return fontAutomation.setProperty(sizePropertyIds[0], new Variant(size));
	}
	
	/**
	 * Make text bold 
	 * @param fontAutomation
	 * @param bold
	 * @return
	 */
	public static boolean setBoldFont(OleAutomation fontAutomation, boolean bold){		
		int[] boldPropertyIds = fontAutomation.getIDsOfNames(new String[]{"Bold"}); 
		return fontAutomation.setProperty(boldPropertyIds[0], new Variant(bold));
	}
	
	/**
	 * Set font color for the text
	 * @param fontAutomation
	 * @param color
	 * @return
	 */
	public static boolean setFontColor(OleAutomation fontAutomation, long color){		
		int[] longPropertyIds = fontAutomation.getIDsOfNames(new String[]{"Color"}); 
		return fontAutomation.setProperty(longPropertyIds[0], new Variant(color));
	}
}
