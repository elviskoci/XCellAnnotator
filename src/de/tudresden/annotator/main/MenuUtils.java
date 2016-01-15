/**
 * 
 */
package de.tudresden.annotator.main;

import org.eclipse.swt.SWT;
import org.eclipse.swt.widgets.Menu;
import org.eclipse.swt.widgets.MenuItem;
import org.eclipse.swt.widgets.MessageBox;

import de.tudresden.annotator.annotations.WorkbookAnnotation;
import de.tudresden.annotator.annotations.WorksheetAnnotation;
import de.tudresden.annotator.annotations.utils.AnnotationDataSheet;
import de.tudresden.annotator.annotations.utils.AnnotationHandler;

/**
 * @author Elvis Koci
 */
public class MenuUtils {
	
	protected static void adjustBarMenuForSheet(){
					
		BarMenu  menuBar = MainWindow.getInstance().getMenuBar();
		MenuItem[] menuItems = menuBar.getMenuItems();
		
		MenuItem annotationsMenu = null;
		for (MenuItem menuItem : menuItems) {
			if(menuItem.getID()==2000000){
				annotationsMenu = menuItem;
				break;
			}
		}
		MenuItem[] annotationsMenuItems = annotationsMenu.getMenu().getItems();
		
		// the active worksheet annotation
		String sheetName = MainWindow.getInstance().getActiveWorksheetName();
		WorkbookAnnotation workbookAnnotation = AnnotationHandler.getWorkbookAnnotation();
		WorksheetAnnotation  sheetAnnotation = workbookAnnotation.getWorksheetAnnotations().get(sheetName);
		
		// if annotation data sheet is the active sheet disable all annotation menus
		// if sheet annotation does not exist, do the same.
		// normally sheet annotation should always exist, but when the annotation data sheet is initially
		// created its name is random. It is after creation that its name is updated to "Annotation_Data_Sheet"
		if(sheetAnnotation==null || sheetName.compareTo(AnnotationDataSheet.getName())==0){
			for (MenuItem menuItem : annotationsMenuItems) {
				menuItem.setEnabled(false);
			}
			return;
		}
		
		// adjust according to the status (Completed, NotApplicable, or non of these) of the sheet. 		
		if(sheetAnnotation.isCompleted()){
		
			for (MenuItem menuItem : annotationsMenuItems) {
				if(menuItem.getID()==2010000 || menuItem.getID()==2050000){ // &Range as and &Delete 
						menuItem.setEnabled(false);
				}else{
						menuItem.setEnabled(true);
						
						if(menuItem.getID()==2020000){ // &Sheet as 
							MenuItem[] submenus = menuItem.getMenu().getItems();
							for (int i = 0; i < submenus.length; i++) {
								if(submenus[i].getID()!=2020200){ // Completed
									submenus[i].setEnabled(false);
								}else{
									submenus[i].setEnabled(true);
								}
							}
						}
				}
			}
			
//			int style = SWT.ICON_INFORMATION;
//			MessageBox mb = MainWindow.getInstance().createMessageBox(style);
//			mb.setMessage("The sheet was marked as \"Completed\""); 
//			mb.open();
			
			
		}else{
			
			if(sheetAnnotation.isNotApplicable()){				
				for (MenuItem menuItem : annotationsMenuItems) { 
					if(menuItem.getID()==2010000 || menuItem.getID()==2040000 || 
					   menuItem.getID()==2050000 || menuItem.getID()==2060000 || 
					   menuItem.getID()==2070000){  
					   // &Range as, &Hide, &Delete, Show Formulas, and Show Annotations 
							menuItem.setEnabled(false);
					}else{
							menuItem.setEnabled(true);
							if(menuItem.getID()==2020000){ // &Sheet as 
								MenuItem[] submenus = menuItem.getMenu().getItems();
								for (int i = 0; i < submenus.length; i++) {
									if(submenus[i].getID()!=2020100){ // Not Applicable
										submenus[i].setEnabled(false);
									}else{
										submenus[i].setEnabled(true);
									}
								}
							}
					}
				}
				
//				int style = SWT.ICON_INFORMATION;
//				MessageBox mb = MainWindow.getInstance().createMessageBox(style);
//				mb.setMessage("The sheet was marked as \"Not Applicable\""); 
//				mb.open();
				
			}else{
				enableAllSubMenus(annotationsMenu.getMenu());				
			}
		}
	}
	
	
	protected static void adjustBarMenuForWorkbook(){
		
		WorkbookAnnotation workbookAnnotation = AnnotationHandler.getWorkbookAnnotation();
		
		BarMenu  menuBar = MainWindow.getInstance().getMenuBar();
		MenuItem[] menuItems = menuBar.getMenuItems();
		
		MenuItem annotationsMenu = null;
		for (MenuItem menuItem : menuItems) {
			if(menuItem.getID()==2000000){ // annotations menu
				annotationsMenu = menuItem;
				break;
			}
		}
		
		if(workbookAnnotation.isCompleted()){
			MenuItem[] annotationsMenuItems = annotationsMenu.getMenu().getItems();
			for (MenuItem menuItem : annotationsMenuItems) { 
				if(menuItem.getID() == 2010000 || menuItem.getID() == 2020000 ||
				   menuItem.getID() == 2050000 || menuItem.getID() == 2060000 ){ 
				   // &Range as, &Sheet as, &Delete, and &Show Formulas
						menuItem.setEnabled(false);
				}else{
						menuItem.setEnabled(true);
						if(menuItem.getID()==2030000){ // &File as 
							MenuItem[] submenus = menuItem.getMenu().getItems();
							for (int i = 0; i < submenus.length; i++) {
								if(submenus[i].getID()!=2030200 ){ // Completed
									submenus[i].setEnabled(false);
								}else{
									submenus[i].setEnabled(true);
								}
							}
						}
				} 
			}			
		}else{
			
			if(workbookAnnotation.isNotApplicable()){
				MenuItem[] annotationsMenuItems = annotationsMenu.getMenu().getItems();
				for (MenuItem menuItem : annotationsMenuItems) {
					if(menuItem.getID() != 2030000){ // &File as
							menuItem.setEnabled(false);
					}else{
							MenuItem[] submenus = menuItem.getMenu().getItems();
							for (int i = 0; i < submenus.length; i++) {
								if(submenus[i].getID()!=2030100 ){ // Not Applicable
									submenus[i].setEnabled(false);
								}else{
									submenus[i].setEnabled(true);
								}
							}
					} 
				}	
			}else{
				adjustBarMenuForSheet();
			}			
		}		
	}
	
	protected static void adjustBarMenuForOpennedFile(){
		
		BarMenu  menuBar = MainWindow.getInstance().getMenuBar();
		MenuItem[] menuItems = menuBar.getMenuItems();
		
		MenuItem fileMenu = null;		
		for (MenuItem menuItem : menuItems) {
			if(menuItem.getID()==1000000){ // file menu
				fileMenu = menuItem;
			}else{
				enableAllSubMenus(menuItem.getMenu());
			}
		}
		
		MenuItem[] fileMenuItems = fileMenu.getMenu().getItems();
		for (MenuItem menuItem : fileMenuItems) {
			if(!(menuItem.getID() == 1020000 || menuItem.getID() == 1030000)){ // Open Prev and Open Next
				menuItem.setEnabled(true);
			}	
		}
		
		adjustBarMenuForWorkbook();
	}
	
	private static void enableAllSubMenus(Menu cascadeMenu){
		if(cascadeMenu!=null){
			MenuItem[]  submenus = cascadeMenu.getItems();
			for (int i = 0; i < submenus.length; i++) {
				enableAllSubMenus(submenus[i].getMenu());
				submenus[i].setEnabled(true);
			}
		}
	}
}
