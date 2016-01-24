/**
 * 
 */
package de.tudresden.annotator.main;

import java.io.BufferedWriter;
import java.io.File;
import java.io.FileWriter;
import java.io.IOException;
import java.nio.file.Files;
import java.util.Collection;

import org.eclipse.swt.SWT;
import org.eclipse.swt.SWTError;
import org.eclipse.swt.ole.win32.OleAutomation;
import org.eclipse.swt.widgets.FileDialog;
import org.eclipse.swt.widgets.MessageBox;

import de.tudresden.annotator.annotations.RangeAnnotation;
import de.tudresden.annotator.annotations.utils.AnnotationHandler;
import de.tudresden.annotator.annotations.utils.AnnotationStatusSheet;
import de.tudresden.annotator.annotations.utils.RangeAnnotationsSheet;
import de.tudresden.annotator.oleutils.ApplicationUtils;
import de.tudresden.annotator.oleutils.WorkbookUtils;

/**
 * @author Elvis Koci
 */
public class FileUtils {
	
	public static final String CurrentProgressFileName = "annotated"; 
	
	public static final String completedFilesFolderName = "completed";
	public static final String notApplicableFilesFolderName = "not_applicable";
	public static final String inProgressFilesFolderName = "in_progress";
	public static final String otherFilesFolderName = "other";
	
	
	/**
	 * Open an excel file for annotation
	 */
	 public static void fileOpen(){
		
		// Select the excel file
		FileDialog dialog = MainWindow.getInstance().createFileDialog(SWT.OPEN);
		String filePath = dialog.open();
		fileOpen(filePath);
	}
	
	 
	/**
	 * Open an excel file for annotation
	 * @param filePath the absolute path of the file to open
	 */
	 public static void fileOpen(String filePath){
		 	 
		// if no file was selected, return
		if (filePath == null) return;
		
		MainWindow mw = MainWindow.getInstance();
		
		// dispose OleControlSite if it is not null. 
		mw.disposeControlSite();
				
	    if (mw.isControlSiteNull()) {
			int index = filePath.lastIndexOf('.');
			if (index != -1) {
				String fileExtension = filePath.substring(index + 1); 
				if (fileExtension.equalsIgnoreCase("xls") || 
						fileExtension.equalsIgnoreCase("xlsx") || 
							fileExtension.equalsIgnoreCase("xlsm")) { // including macro enabled ?	
					
					try {		    	
				        
						File excelFile = new File(filePath);
				        
				        // embed the excel file and set up the user interface
				        mw.embedExcelFile(excelFile);
				        
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
	 
	 
	/**
	 * Save all the annotation progress 
	 * @param embeddedWorkbook an OleAutomation that provides access to the functionalities of the embedded workbook
	 * @param filePath the path where to save the file
	 * @param beforeFileClose true if progress is saved before closing the file or exiting the application, 
	 * false if file will remain open after save.
	 * @return true if progress was successfully saved, false otherwise. 
	 */
	public static boolean saveProgress(OleAutomation embeddedWorkbook, String filePath, boolean beforeFileClose){
				 
		// save the status of all worksheet annotations and the workbook annotation 
		AnnotationStatusSheet.saveAnnotationStatuses(embeddedWorkbook);
		
		// delete all shape annotations
		AnnotationHandler.deleteAllShapeAnnotations(embeddedWorkbook);
						
		// unprotect the workbook structure and all the sheets
		WorkbookUtils.unprotectWorkbook(embeddedWorkbook);
		WorkbookUtils.unprotectAllWorksheets(embeddedWorkbook);

		// protect and hide the range_annotations sheet before save
		RangeAnnotationsSheet.protect(embeddedWorkbook);
		RangeAnnotationsSheet.setVisibility(embeddedWorkbook, false);
		
		// protect and hide the annotation_status sheet before save
		AnnotationStatusSheet.protect(embeddedWorkbook);
		AnnotationStatusSheet.setVisibility(embeddedWorkbook, false);
		
		// deactivate alerts before save
		OleAutomation application = WorkbookUtils.getApplicationAutomation(embeddedWorkbook);
		MainWindow.getInstance().deactivateControlSite();
		ApplicationUtils.setDisplayAlerts(application, false);
		
		// save the file
	
		boolean isSuccess = WorkbookUtils.saveWorkbookAs(embeddedWorkbook, filePath, null);
		WorkbookUtils.closeEmbeddedWorkbook(embeddedWorkbook, false);
		MainWindow.getInstance().setEmbeddedWorkbook(null);
	
		// activate alerts after save
		ApplicationUtils.setDisplayAlerts(application, true);
		
		String newPath =  moveFileToStatusDirectory();
			
		if(!beforeFileClose){			
			//MainWindow.getInstance().doVerbControlSite(OLE.OLEIVERB_INPLACEACTIVATE);
			//MainWindow.getInstance().setUpApplicationDisplay(application);
			
			fileOpen(newPath);
					
			OleAutomation reopenedWorkbook = MainWindow.getInstance().getEmbeddedWorkbook();
			
			// draw again the range annotations  
			Collection<RangeAnnotation> collection= AnnotationHandler.getWorkbookAnnotation().getAllAnnotations();
			RangeAnnotation[] rangeAnnotations = collection.toArray(new RangeAnnotation[collection.size()]);
			if(rangeAnnotations!=null){		
				
				// update workbook annotation and re-draw all the range annotations  
				AnnotationHandler.drawManyRangeAnnotations(reopenedWorkbook, rangeAnnotations);	
			}
						
			// make range_annotations sheet again visible
			RangeAnnotationsSheet.setVisibility(reopenedWorkbook, true);
		}
		return isSuccess;
	}


	/**
	 * Move the opened (embedded) excel file to the directory that corresponds to 
	 * its current annotation status. For example, if the file was marked as "Completed",
	 * it will be moved to the folder where all the completed files are grouped (placed).
	 * 
	 */
	public static String moveFileToStatusDirectory(){
		
		String fileName = MainWindow.getInstance().getFileName();
		String fileDirPath = MainWindow.getInstance().getDirectoryPath();
		
		File file = new File(fileDirPath+"\\"+fileName);
		File directory = new File(fileDirPath); 
				
		String originalDir = fileDirPath; 
		
		if(directory.getName().compareTo(completedFilesFolderName)==0 || 
		   directory.getName().compareTo(notApplicableFilesFolderName)==0 || 
		   directory.getName().compareTo(inProgressFilesFolderName)==0 ){
				
			originalDir = directory.getParentFile().getAbsolutePath();
		}
		
		File newLocation = file;
		
		if(AnnotationHandler.getWorkbookAnnotation().isCompleted()){
			
			File completed = new File (originalDir+"\\"+completedFilesFolderName);
			if(!completed.exists())
				completed.mkdir();
			
			newLocation = new File(completed.getAbsolutePath()+"\\"+fileName); 
				
			if(!file.getParent().equals(completed))
				moveFile(file, newLocation);
					
		}else if(AnnotationHandler.getWorkbookAnnotation().isNotApplicable()){
			
			File notApplicable = new File (originalDir+"\\"+notApplicableFilesFolderName);
			if(!notApplicable.exists())
				notApplicable.mkdir();
			
			newLocation = new File(notApplicable.getAbsolutePath()+"\\"+fileName); 
			
			if(!file.getParent().equals(notApplicable))
				moveFile(file, newLocation);
			
		}else{
			
			File inProgress = new File (originalDir+"\\"+inProgressFilesFolderName);
			if(!inProgress.exists())
				inProgress.mkdir();
			
			newLocation = new File(inProgress.getAbsolutePath()+"\\"+fileName); 
			
			if(!file.getParent().equals(inProgress))
				moveFile(file, newLocation);			
		}
		
		return newLocation.getAbsolutePath();
	}
	
	
	/**
	 * 
	 * @param file
	 * @param directory
	 */
	private static void moveFile(File file, File directory){
		
		try {
			Files.move(file.toPath(), directory.toPath());
		} catch (IOException e) {
			
			MessageBox message = MainWindow.getInstance().createMessageBox(SWT.ICON_ERROR);
			message.setText("ERROR");
			message.setMessage("ERROR: Could not move file \""+file.getName()+"\" "
					+ "to the directory \""+directory.getName()+"\". \n\n"
							+ e.toString());
			message.open();
		}
		
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
}
