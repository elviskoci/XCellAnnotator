/**
 * 
 */
package de.tudresden.annotator.oleutils;

import java.io.BufferedWriter;
import java.io.File;
import java.io.FileWriter;
import java.io.IOException;

import org.eclipse.swt.ole.win32.OleAutomation;
import org.eclipse.swt.ole.win32.OleClientSite;
import org.eclipse.swt.ole.win32.OleControlSite;
import org.eclipse.swt.ole.win32.Variant;

import de.tudresden.annotator.annotations.utils.AnnotationDataSheet;
import de.tudresden.annotator.annotations.utils.AnnotationHandler;
import de.tudresden.annotator.main.MainWindow;

/**
 * @author Elvis Koci
 */
public class FileUtils {
	
	public static final String CurrentProgressFileName = "annotated"; 
	
	public static boolean saveProgress(OleAutomation embeddedWorkbook, String filePath){
		
		boolean isSaved = WorkbookUtils.isWorkbookSaved(embeddedWorkbook);
		if(isSaved){
			System.out.println("There are no changes since last save");
			return false;
		}
			
		boolean isUnprotected = WorkbookUtils.unprotectWorkbook(embeddedWorkbook);		
		if(!isUnprotected){
			System.out.println("ERROR: Could not unprotect the workbook. Operation failed!");
			return false;
			//System.exit(1);
		}

		boolean areUnprotected =WorkbookUtils.unprotectAllWorksheets(embeddedWorkbook);
		if(!areUnprotected){
			System.out.println("ERROR: Could not unprotect all worksheets. Operation failed!");
			return false;
			//System.exit(1);
		}
		
		AnnotationHandler.setVisilityForShapeAnnotations(embeddedWorkbook, false);
		AnnotationDataSheet.setVisibility(embeddedWorkbook, false);
		
		System.out.println("The file path is: "+filePath);
		boolean isSuccess = WorkbookUtils.saveWorkbookAs(embeddedWorkbook, filePath, null);
		
		AnnotationDataSheet.setVisibility(embeddedWorkbook, true);
		AnnotationHandler.setVisilityForShapeAnnotations(embeddedWorkbook, true);
		
		boolean isProtected = WorkbookUtils.protectWorkbook(embeddedWorkbook, true, false);		
		if(!isProtected){
			System.out.println("ERROR: Could not protect the workbook. Operation failed!");
			return false;
			//System.exit(1);
		}
		
		boolean areProtected = WorkbookUtils.protectAllWorksheets(embeddedWorkbook);
		if(!areProtected){
			System.out.println("ERROR: Could not protect all worksheets. Operation failed!");
			return false;
			//System.exit(1);
		}
		
		return isSuccess;
	}

	
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
