/**
 * 
 */
package de.tudresden.annotator.main;

import org.eclipse.swt.widgets.Menu;
import org.eclipse.swt.widgets.MenuItem;

import de.tudresden.annotator.annotations.WorkbookAnnotation;
import de.tudresden.annotator.annotations.WorksheetAnnotation;
import de.tudresden.annotator.annotations.utils.AnnotationHandler;
import de.tudresden.annotator.annotations.utils.RangeAnnotationsSheet;

/**
 * @author Elvis Koci
 */
public class MenuUtils {
	
	protected static void adjustBarMenuForSheet(String sheetName){
					
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
		//String sheetName = MainWindow.getInstance().getActiveWorksheetName();
		WorkbookAnnotation workbookAnnotation = AnnotationHandler.getWorkbookAnnotation();
		WorksheetAnnotation  sheetAnnotation = workbookAnnotation.getWorksheetAnnotations().get(sheetName);
		
		
		// if annotation data sheet is the active sheet disable all annotation menus
		// if sheet annotation does not exist, do the same.
		// normally sheet annotation should always exist, but when the range annotations sheet is initially
		// created its name is random. It is after creation that its name is updated to "Range_Annotations_Data"
		if(sheetAnnotation==null || sheetName.compareTo(RangeAnnotationsSheet.getName())==0){
			for (MenuItem menuItem : annotationsMenuItems) {
				menuItem.setEnabled(false);
			}
		}else{
			
			// adjust according to the status (Completed, NotApplicable, or non of these) of the sheet. 		
			if(sheetAnnotation.isCompleted()){
			
				for (MenuItem menuItem : annotationsMenuItems) {
					
					if(menuItem.getID() == 2010000 || menuItem.getID() == 2050000 ||
					   menuItem.getID() == 2080000 || menuItem.getID() == 2090000){ 
							// &Range as, &Delete, Undo, and Redo 
							menuItem.setEnabled(false);
					}else{ 
							menuItem.setEnabled(true);
							if(menuItem.getID()==2020000){ // &Sheet as 
							
								MenuItem[] submenus = menuItem.getMenu().getItems();
								for (int i = 0; i < submenus.length; i++) {
									if(submenus[i].getID()==2020100){ // Not Applicable
										submenus[i].setEnabled(false);
										submenus[i].setSelection(false);
									}
									
									if(submenus[i].getID()==2020200){ // Completed
										submenus[i].setEnabled(true);
										submenus[i].setSelection(true);
									}
								}
							}
					}
				}
							
			}else{
				
				if(sheetAnnotation.isNotApplicable()){				
					for (MenuItem menuItem : annotationsMenuItems) { 
						if(menuItem.getID()==2010000 || menuItem.getID()==2040000 || 
						   menuItem.getID()==2050000 || menuItem.getID()==2060000 || 
						   menuItem.getID()==2070000 || menuItem.getID() == 2080000 || 
						   menuItem.getID() == 2090000){  
						   // &Range as, &Hide, &Delete, Show Formulas, and Show Annotations 
								menuItem.setEnabled(false);
						}else{
								menuItem.setEnabled(true);
								if(menuItem.getID()==2020000){ // &Sheet as 
									MenuItem[] submenus = menuItem.getMenu().getItems();
									for (int i = 0; i < submenus.length; i++) {
										if(submenus[i].getID()==2020100){ // Not Applicable
											submenus[i].setEnabled(true);
											submenus[i].setSelection(true);
										}
										
										if(submenus[i].getID()==2020200){ // Completed 
											submenus[i].setEnabled(false);
											submenus[i].setSelection(false);
										}
									}
								}
						}
					}
					
				}else{
					
					boolean hasAnnotations = sheetAnnotation.getAllAnnotations().size()>0;		
					if(!hasAnnotations){
						for (MenuItem menuItem : annotationsMenuItems) { 
							if(menuItem.getID() == 2040000 || menuItem.getID() == 2050000 || 
							   menuItem.getID() == 2060000 || menuItem.getID() == 2070000 ||
							   menuItem.getID() == 2080000 ){
							   // &Hide, &Delete, Show Formulas, &Show 
							   menuItem.setEnabled(false);
							}else if(menuItem.getID() == 2090000){ // Redo last annotation
								if(AnnotationHandler.getLastFromRedoList()==null){
										menuItem.setEnabled(false);
								}else{
										menuItem.setEnabled(true);
								}
								
							}else{
								menuItem.setEnabled(true);
								enableAllSubMenus(menuItem.getMenu());
								unselectAllSubMenus(menuItem.getMenu());
							}
						}
					}else{
						for (MenuItem menuItem : annotationsMenuItems) { 
							if( menuItem.getID() == 2080000){	// Undo last annotations				
								if(AnnotationHandler.getLastFromUndoList()==null){
									menuItem.setEnabled(false);
								}else{
									menuItem.setEnabled(true);
								}
								
							}else if(menuItem.getID() == 2090000){ // Redo last annotation
								
								if(AnnotationHandler.getLastFromRedoList()==null){
									menuItem.setEnabled(false);
								}else{
									menuItem.setEnabled(true);
								}
							
							}else{
								menuItem.setEnabled(true);
								enableAllSubMenus(menuItem.getMenu());
								unselectAllSubMenus(menuItem.getMenu());
							}
						}
					}									
				}
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
				   menuItem.getID() == 2050000 || menuItem.getID() == 2060000 ||
				   menuItem.getID() == 2080000 || menuItem.getID() == 2090000 ){ 
				   // &Range as, &Sheet as, &Delete,&Show Formulas, Undo and Redo
						menuItem.setEnabled(false);
				}else{
						menuItem.setEnabled(true);
						if(menuItem.getID()==2030000){ // &File as 
							MenuItem[] submenus = menuItem.getMenu().getItems();
							for (int i = 0; i < submenus.length; i++) {
								
								if(submenus[i].getID()==2030100){ // Not Applicable
									submenus[i].setEnabled(false);
									submenus[i].setSelection(false);
								}
								
								if(submenus[i].getID()==2030200){ // Completed 
									submenus[i].setEnabled(true);
									submenus[i].setSelection(true);
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
							for (int i = 0; i < submenus.length; i++){
								if(submenus[i].getID()==2030100){ // Not Applicable
									submenus[i].setEnabled(true);
									submenus[i].setSelection(true);
								}
								
								if(submenus[i].getID()==2030200){ // Completed 
									submenus[i].setEnabled(false);
									submenus[i].setSelection(false);
								}
							}
					} 
				}	
			}else{

				MenuItem[] annotationsMenuItems = annotationsMenu.getMenu().getItems();
				for (MenuItem menuItem : annotationsMenuItems) {
					if(menuItem.getID() == 2030000){ // &File as
						MenuItem[] submenus = menuItem.getMenu().getItems();
						for (int i = 0; i < submenus.length; i++) {
							if(submenus[i].getID()==2030100){ // Not Applicable
								submenus[i].setEnabled(true);
								submenus[i].setSelection(false);
							}
							
							if(submenus[i].getID()==2030200){ // Completed 
								submenus[i].setEnabled(true);
								submenus[i].setSelection(false);
							}
						}
					}
				}			
						
				adjustBarMenuForSheet(MainWindow.getInstance().getActiveWorksheetName());
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
				// unselectAllSubMenus(menuItem.getMenu());
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
	
	private static void unselectAllSubMenus(Menu cascadeMenu){
		if(cascadeMenu!=null){
			MenuItem[]  submenus = cascadeMenu.getItems();
			for (int i = 0; i < submenus.length; i++) {
				enableAllSubMenus(submenus[i].getMenu());
				submenus[i].setSelection(false);
			}
		}
	}
}
