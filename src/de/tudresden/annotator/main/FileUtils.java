/**
 * 
 */
package de.tudresden.annotator.main;

import java.io.BufferedWriter;
import java.io.File;
import java.io.FileWriter;
import java.io.IOException;

import org.eclipse.swt.SWT;
import org.eclipse.swt.SWTError;
import org.eclipse.swt.ole.win32.OleAutomation;
import org.eclipse.swt.widgets.FileDialog;
import org.eclipse.swt.widgets.MessageBox;

import de.tudresden.annotator.annotations.utils.AnnotationDataSheet;
import de.tudresden.annotator.annotations.utils.AnnotationHandler;
import de.tudresden.annotator.oleutils.WorkbookUtils;

/**
 * @author Elvis Koci
 */
public class FileUtils {
	
	public static final String CurrentProgressFileName = "annotated"; 
	
	/**
	 * 
	 * @param embeddedWorkbook
	 * @param filePath
	 * @return
	 */
	public static boolean saveProgress(OleAutomation embeddedWorkbook, String filePath){
		
		boolean isSaved = WorkbookUtils.isWorkbookSaved(embeddedWorkbook);
		if(isSaved){
			System.out.println("There are no changes since last save");
			return false;
		}
		
		//delete all shape annotations
		AnnotationHandler.deleteAllShapeAnnotations(embeddedWorkbook);
	
		boolean isUnprotected = WorkbookUtils.unprotectWorkbook(embeddedWorkbook);		
		if(!isUnprotected){
			System.out.println("ERROR: Could not unprotect the workbook. Operation failed!");
			return false;
		}

		boolean areUnprotected =WorkbookUtils.unprotectAllWorksheets(embeddedWorkbook);
		if(!areUnprotected){
			System.out.println("ERROR: Could not unprotect all worksheets. Operation failed!");
			return false;
		}
		
		AnnotationDataSheet.protect(embeddedWorkbook);
		AnnotationDataSheet.setVisibility(embeddedWorkbook, false);
		
		boolean isSuccess = WorkbookUtils.saveWorkbookAs(embeddedWorkbook, filePath, null);
		
		AnnotationDataSheet.setVisibility(embeddedWorkbook, true);
		// draw again the range annotations  
		AnnotationHandler.drawAllAnnotations(embeddedWorkbook);

		boolean isProtected = WorkbookUtils.protectWorkbook(embeddedWorkbook, true, false);		
		if(!isProtected){
			System.out.println("ERROR: Could not protect the workbook. Operation failed!");
			return false;
		}
		
		boolean areProtected = WorkbookUtils.protectAllWorksheets(embeddedWorkbook);
		if(!areProtected){
			System.out.println("ERROR: Could not protect all worksheets. Operation failed!");
			return false;
		}
		
		return isSuccess;
	}

	/**
	 * 
	 * @param directory
	 * @param fileName
	 * @param status
	 * @return
	 */
	public static boolean markFileAsAnnotated(String directory, String fileName, int status){
		
		File file = new File(directory+"\\"+CurrentProgressFileName);
		
		try {
			if (!file.exists()) {
				file.createNewFile();
			}
			
			FileWriter fw = new FileWriter(file.getAbsoluteFile());
			BufferedWriter bw = new BufferedWriter(fw);
			
			String content = fileName+"\t"+status+"\n";
			bw.write(content);
			bw.close();
			
		} catch (IOException e) {
			e.printStackTrace();
			return false;
		}
		
		return true;
	}
	
	
	/**
	 * Open an excel file for annotation
	 */
	 public static void fileOpen(){
		
		MainWindow mw = MainWindow.getInstance();
		
		// Select the excel file
		FileDialog dialog = mw.createFileDialog(SWT.OPEN);
		String fileName = dialog.open();
		
		// if no file was selected, return
		if (fileName == null) return;
		
		// dispose OleControlSite if it is not null. 
		mw.disposeControlSite();
				
	    if (mw.isControlSiteNull()) {
			int index = fileName.lastIndexOf('.');
			if (index != -1) {
				String fileExtension = fileName.substring(index + 1); 
				if (fileExtension.equalsIgnoreCase("xls") || fileExtension.equalsIgnoreCase("xlsx") || fileExtension.equalsIgnoreCase("xlsm")) { // including macro enabled ?	
					
					try {		    	
				        
						File excelFile = new File(fileName);
				        
				        // set up the excel application user interface for the annotation task
				        mw.setUpWorkbookDisplay(excelFile);
				        
				    } catch (SWTError e) {
				        e.printStackTrace();
				        System.out.println("Unable to open ActiveX Control");
				        return;
				    }	    	  
				   
				}else{
					MessageBox msgbox = mw.createMessageBox(SWT.ICON_ERROR);
					msgbox.setMessage("The selected file format is not recognized: ."+fileExtension);
					msgbox.open();
				}
			}
	    }
	}
}
