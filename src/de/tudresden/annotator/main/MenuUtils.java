/**
 * 
 */
package de.tudresden.annotator.main;

import org.eclipse.swt.widgets.MenuItem;

import de.tudresden.annotator.annotations.WorkbookAnnotation;
import de.tudresden.annotator.annotations.WorksheetAnnotation;
import de.tudresden.annotator.annotations.utils.AnnotationHandler;

/**
 * @author Elvis Koci
 */
public class MenuUtils {
	
	public static void adjustMenuForSheet(){
		
		String sheetName = MainWindow.getInstance().getActiveWorksheetName();
		WorkbookAnnotation workbookAnnotation = AnnotationHandler.getWorkbookAnnotation();
		WorksheetAnnotation sheetAnnotation = workbookAnnotation.getWorksheetAnnotations().get(sheetName);
		
//		if(sheetAnnotation.isCompleted()){
//			sheetAnnotation.setCompleted(false);
//			
//			MenuItem annotationsMenu = menuItems[1]; // Annotations menu 
//			MenuItem[] annotationsSubmenus = annotationsMenu.getMenu().getItems();
//			for (MenuItem menuItem : annotationsSubmenus) {
//					menuItem.setEnabled(true);
//			}
//			annotateWorksheetMenuItem.getMenu().getItem(0).setEnabled(true); // enable "NotApplicable" menu item.					
//			strMessage = "The status for the sheet \""+sheetName+"\" was changed back to \"Not Completed\"";
//			
//		}else{
//			
//			sheetAnnotation.setCompleted(true);
//			
//			MenuItem annotationsMenu = menuItems[1]; // Annotations menu 
//			MenuItem[] annotationsSubmenus = annotationsMenu.getMenu().getItems();
//			for (MenuItem menuItem : annotationsSubmenus) {
//				if(menuItem.getText().compareTo("&Range as")==0 || menuItem.getText().compareTo("&Delete")==0){
//					menuItem.setEnabled(false);
//				}else{
//					menuItem.setEnabled(true);
//				}
//			}
//			annotateWorksheetMenuItem.getMenu().getItem(0).setEnabled(false); // disable "NotApplicable" menu item.
//			
//			strMessage = "The sheet \""+sheetName+"\" was marked completed";
//		}
	}

}
