package de.tudresden.annotator.main;

import java.util.ArrayList;
import java.util.Arrays;
import java.util.Iterator;
import java.util.LinkedHashMap;

import org.eclipse.swt.SWT;
import org.eclipse.swt.widgets.Menu;
import org.eclipse.swt.widgets.MenuItem;
import org.eclipse.swt.widgets.Shell;

import de.tudresden.annotator.annotations.AnnotationClass;
import de.tudresden.annotator.annotations.utils.ClassGenerator;

public class BarMenu {
	
	private MenuItem[] menuItems;
	
	public BarMenu(Shell oleShell){
		
		final Shell shell = oleShell;
		
		Menu menuBar = shell.getMenuBar();
		if (menuBar == null) {
			menuBar = new Menu(shell, SWT.BAR);
			shell.setMenuBar(menuBar);
		}
	
		menuItems = new MenuItem[4];
		menuItems[0] = addFileMenu(menuBar);
		menuItems[1] = addAnnotationsMenu(menuBar);
		menuItems[2] = addViewMenu(menuBar);
		menuItems[3] = addPreferencesMenu(menuBar);
		
		setToolTipText(menuBar);
	}
	
	
	private  void setToolTipText(Menu cascadeMenu){
		if(cascadeMenu!=null){
			for (int i = 0; i < cascadeMenu.getItems().length; i++) {
				setToolTipText(cascadeMenu.getItem(i).getMenu());
				cascadeMenu.getItem(i).setToolTipText(""+cascadeMenu.getItem(i).getID());
			}			
		}else{
			return;
		}		
	}
	
	/**
	 * Create the "File" menu
	 * @param menuBar the (parent) bar menu 
	 * @return a menu item containing the "File" menu options
	 */
	private MenuItem addFileMenu(Menu menuBar){	
		
		MenuItem fileMenu = new MenuItem(menuBar, SWT.CASCADE);
		fileMenu.setID(1000000);
		fileMenu.setText("&File");
		Menu menuFile = new Menu(fileMenu);
		fileMenu.setMenu(menuFile);
		
		/*
		 *  Open File menu item
		 */
		MenuItem menuFileOpen = new MenuItem(menuFile, SWT.CASCADE);
		menuFileOpen.setID(1010000);
		menuFileOpen.setText("Open... \tCtrl+O");
		menuFileOpen.setAccelerator(SWT.MOD1+'O');
		menuFileOpen.addSelectionListener(GUIListeners.createFileOpenSelectionListener());
		
		/*
		 *  Open Previous File menu item 
		 */
		MenuItem menuFileOpenPrevious = new MenuItem(menuFile, SWT.CASCADE);
		menuFileOpenPrevious.setID(1020000);
		menuFileOpenPrevious.setText("Open Prev");
		menuFileOpenPrevious.setEnabled(false);
	
		
		/*
		 *  Open Next File menu item 
		 */
		MenuItem menuFileOpenNext = new MenuItem(menuFile, SWT.CASCADE);
		menuFileOpenNext.setID(1030000);
		menuFileOpenNext.setText("Open Next");
		menuFileOpenNext.setEnabled(false);
		
		/*
		 *  Save File menu item 
		 */
		MenuItem menuFileSave = new MenuItem(menuFile, SWT.CASCADE);
		menuFileSave.setID(1040000);
		menuFileSave.setText("Save \tCtrl+S");
		menuFileSave.setEnabled(false);
		menuFileSave.setAccelerator(SWT.MOD1 + 'S');
		
		/*
		 *  Export File menu item 
		 */
		MenuItem menuFileExport = addExportMenu(menuFile);
		menuFileExport.setID(1050000);
		menuFileExport.setEnabled(false);
	
		/*
		 *  Close File menu item 
		 */
		MenuItem menuFileClose = new MenuItem(menuFile, SWT.CASCADE);
		menuFileClose.setID(1060000);
		menuFileClose.setText("Close");
		menuFileClose.setEnabled(false);
	
		/*
		 *  Exit Application menu item 
		 */
		MenuItem menuFileExit = new MenuItem(menuFile, SWT.CASCADE);
		menuFileExit.setID(1070000);
		menuFileExit.setText("Exit \tCtrl+Q");
		menuFileExit.setAccelerator(SWT.MOD1 + 'Q');
		
		return fileMenu;
	}
	
	
	/**
	 * 
	 * @param menuBar
	 * @return
	 */
	private MenuItem addAnnotationsMenu(Menu menuBar){
		
		MenuItem annotationsMenu = new MenuItem(menuBar, SWT.CASCADE);
		annotationsMenu.setText("&Annotations");
		Menu menuAnnotations = new Menu(annotationsMenu);
		annotationsMenu.setMenu(menuAnnotations);
		annotationsMenu.setID(2000000);
		
		/*
		 * Range annotations menu item  
		 */
		MenuItem menuItemRange = addAnnotateRangeMenu(menuAnnotations);
		menuItemRange.setEnabled(false);
		
		/*
		 * Worksheet annotations menu item  
		 */
		MenuItem menuItemSheet = addAnnotateWorksheetMenu(menuAnnotations);	
		menuItemSheet.setEnabled(false);
		
		/*
		 * Workbook annotations menu item  
		 */
		MenuItem menuItemBook = addAnnotateWorkbookMenu(menuAnnotations);
		menuItemBook.setEnabled(false);
		
		/*
		 * Hide annotations menu item  
		 */
		MenuItem menuItemHide = addHideMenu(menuAnnotations);
		menuItemHide.setEnabled(false);
		
		/*
		 * Delete annotations menu item  
		 */
		MenuItem menuItemDelete = addDeleteMenu(menuAnnotations);
		menuItemDelete.setEnabled(false);
		
		/*
		 * Show Formulas menu item  
		 */
		MenuItem menuItemShowFormulas = new MenuItem(menuAnnotations, SWT.CASCADE);
		menuItemShowFormulas.setText("Show Formulas");
		menuItemShowFormulas.setEnabled(false);
		menuItemShowFormulas.setID(2060000);
		
		/*
		 * Show Annotations menu item  
		 */
		MenuItem menuItemShowAnnotations = new MenuItem(menuAnnotations, SWT.CASCADE);
		menuItemShowAnnotations.setText("Show Annotations");
		menuItemShowAnnotations.setEnabled(false);
		menuItemShowAnnotations.setID(2070000);
		
		return annotationsMenu;
	}
	
	private MenuItem addViewMenu(Menu menuBar) {
		
		MenuItem viewMenu = new MenuItem(menuBar, SWT.CASCADE);
		viewMenu.setText("&View");
		Menu menuView = new Menu(viewMenu);
		viewMenu.setMenu(menuView);
		viewMenu.setID(3000000);
		
		/*
		 * View Folder Explorer Panel menu item  
		 */
		MenuItem menuViewFolderExplorer = new MenuItem(menuView, SWT.CASCADE);
		menuViewFolderExplorer.setText("Folder Explorer");
		menuViewFolderExplorer.setID(3010000);
		
		/*
		 * View Annotation Management Panel menu item  
		 */
		MenuItem menuViewAnnotationsPanel = new MenuItem(menuView, SWT.CASCADE);
		menuViewAnnotationsPanel.setText("Annotations Panel");
		menuViewAnnotationsPanel.setID(3020000);
		
		return viewMenu;
	}
	
	private MenuItem addPreferencesMenu(Menu menuBar){
		
		MenuItem preferencesMenu = new MenuItem(menuBar, SWT.CASCADE);
		preferencesMenu.setText("&Preferences");
		Menu menuPreferences = new Menu(preferencesMenu);
		preferencesMenu.setMenu(menuPreferences);
		preferencesMenu.setID(4000000);
		
		/*
		 * Formatting Preferences menu item 
		 */
		MenuItem menuPreferencesFormatting = new MenuItem(menuPreferences, SWT.CASCADE);
		menuPreferencesFormatting.setText("Formatting");
		menuPreferencesFormatting.setID(4010000);
		
		/*
		 * Annotation Classes Preferences menu item 
		 */
		MenuItem menuPreferencesAnnotationClasses = new MenuItem(menuPreferences, SWT.CASCADE);
		menuPreferencesAnnotationClasses.setText("Annotation Classes");
		menuPreferencesAnnotationClasses.setID(4020000);
		
		return preferencesMenu;
	}

	
	private MenuItem addAnnotateRangeMenu(Menu menu){
		
		MenuItem annotateRangeMenuItem = new MenuItem(menu, SWT.CASCADE);
		annotateRangeMenuItem.setText("&Range as");
		Menu menuAnnotateRange = new Menu(annotateRangeMenuItem);
		annotateRangeMenuItem.setMenu(menuAnnotateRange);
		annotateRangeMenuItem.setID(2010000);
		
		// using the leftmost characters to make it easier for the user to simultaneously handle the mouse and keyboard
		LinkedHashMap<String, AnnotationClass> map =  ClassGenerator.getAnnotationClasses();
		ArrayList<Character> usedChars = new ArrayList<Character>(); 
		ArrayList<Character> shortcutChars  = new ArrayList<Character>();
		
		Iterator<String> keys = map.keySet().iterator();
		while(keys.hasNext()){
			String label = keys.next();
			if(!usedChars.contains(label.charAt(0))){
				usedChars.add(label.charAt(0));
				shortcutChars.add(label.charAt(0));
			}else{
				shortcutChars.add('?');
			}
		}
		
//		Iterator<String> keys = map.keySet().iterator();
//		
//		ArrayList<Character> shortcutChars  = new ArrayList<Character>();
//		shortcutChars.addAll(Arrays.asList(new Character[]{'A', 'S', 'D', 'X', 'Z', 'C', 'Q', 'W', 'E'}));
//		ArrayList<Character> usedChars = new ArrayList<Character>(); 
//		
//		/* if the label of the annotation class starts with one of the characters 
//		 * in the shortcutChars list, use that character for creating the shortcut of this class
//		 * for classes who's name does not start with one of the considered characters
//		 * use the next available character in the list
//		 */
//		while(keys.hasNext()){
//			String label = keys.next();
//			if(shortcutChars.contains(label.charAt(0))){
//				AnnotationClass ac = map.get(label);
//				ac.setShortcut(SWT.MOD1 | SWT.MOD2 + label.charAt(0));
//				usedChars.add(label.charAt(0));
//			}
//		}
//		
//		for (int i = 0; i < usedChars.size(); i++) {
//			shortcutChars.remove(usedChars.get(i));
//		}
		
		/*
		 * Iterate through the list of annotation classes, and
		 * for each one create a menu item based on its properties
		 */
		Iterator<AnnotationClass> values = map.values().iterator();
		int i = 0;
		while(values.hasNext()){
			
			AnnotationClass annotationClass = (AnnotationClass) values.next();
			MenuItem menuAnnotationClass = new MenuItem(menuAnnotateRange, SWT.CASCADE);
			
			menuAnnotationClass.setID(annotateRangeMenuItem.getID()+((i+1)*100));			
			menuAnnotationClass.addSelectionListener(
					GUIListeners.createRangeAnnotationSelectionListener(annotationClass));
			
			int shortcut = annotationClass.getShortcut();	
			char c = 'A';
			if(shortcut < 0){
				if(shortcutChars.get(i)=='?'){
					while(true){
						if (!usedChars.contains(c)) {
							usedChars.add(c);
							shortcutChars.set(i, c);
							c++;
							break;
						}
						c++;
					}
				}
				shortcut =  shortcutChars.get(i); // SWT.MOD1 | SWT.MOD2 + shortcutChars.get(i);
				annotationClass.setShortcut(shortcut);
			}
			menuAnnotationClass.setAccelerator(shortcut);
			menuAnnotationClass.setText(annotationClass.getLabel()+"\t"+shortcutChars.get(i)); //Ctrl+Shift+

			i++;
		}	
		return annotateRangeMenuItem;
	}
	
	private MenuItem addAnnotateWorksheetMenu(Menu menu){
		
		MenuItem annotateWorksheetMenuItem = new MenuItem(menu, SWT.CASCADE);
		annotateWorksheetMenuItem.setID(2020000);
		annotateWorksheetMenuItem.setText("&Sheet as");
		Menu menuAnnotateWorksheet = new Menu(annotateWorksheetMenuItem);
		annotateWorksheetMenuItem.setMenu(menuAnnotateWorksheet);
		
		/*
		 * Worksheet Not Applicable menu item
		 */
		MenuItem menuItemNotApplicable = new MenuItem(menuAnnotateWorksheet, SWT.CASCADE);
		menuItemNotApplicable.setID(2020100);
		menuItemNotApplicable.setText("Not Applicable");
		menuItemNotApplicable.addSelectionListener(GUIListeners.createSheetNotApplicableSelectionListener());
		
		/*
		 * Worksheet Completed menu item
		 */
		MenuItem menuItemCompleted = new MenuItem(menuAnnotateWorksheet, SWT.CASCADE);
		menuItemCompleted.setID(2020200);
		menuItemCompleted.setText("Completed");
		menuItemCompleted.addSelectionListener(GUIListeners.createSheetCompletedSelectionListener());
		
		return annotateWorksheetMenuItem;
	}
	
	
	private MenuItem addAnnotateWorkbookMenu(Menu menu){
		
		MenuItem annotateWorkbookMenuItem = new MenuItem(menu, SWT.CASCADE);
		annotateWorkbookMenuItem.setID(2030000);
		annotateWorkbookMenuItem.setText("&File as");
		Menu menuAnnotateWorkbook = new Menu(annotateWorkbookMenuItem);
		annotateWorkbookMenuItem.setMenu(menuAnnotateWorkbook);
		
		/*
		 * Workbook Not Applicable menu item
		 */
		MenuItem menuItemNotApplicable = new MenuItem(menuAnnotateWorkbook, SWT.CASCADE);
		menuItemNotApplicable.setID(2030100);
		menuItemNotApplicable.setText("Not Applicable");
		menuItemNotApplicable.addSelectionListener(GUIListeners.createFileNotApplicableSelectionListener());
		
		
		/*
		 * Workbook Completed menu item 
		 */
		MenuItem menuItemCompleted = new MenuItem(menuAnnotateWorkbook, SWT.CASCADE);
		menuItemCompleted.setID(2030200);
		menuItemCompleted.setText("Completed");
		menuItemCompleted.addSelectionListener(GUIListeners.createFileCompletedSelectionListener());
		
		return annotateWorkbookMenuItem;
	}
	
	private MenuItem addHideMenu(Menu menu){
		
		MenuItem hideMenuItem = new MenuItem(menu, SWT.CASCADE);
		hideMenuItem.setText("&Hide");
		Menu menuHide = new Menu(hideMenuItem);
		hideMenuItem.setMenu(menuHide);
		hideMenuItem.setID(2040000);
		
		/*
		 * Hide All menu item
		 */
		MenuItem menuItemClearAll = new MenuItem(menuHide, SWT.CASCADE);
		menuItemClearAll.setText("All");
		menuItemClearAll.setID(2040100);
		
		/*
		 * Hide In Sheet menu item
		 */
		MenuItem menuItemClearInSheet = new MenuItem(menuHide, SWT.CASCADE);
		menuItemClearInSheet.setText("In Sheet");
		menuItemClearInSheet.setID(2040200);
	
		return hideMenuItem;
	}
	
	
	private MenuItem addDeleteMenu(Menu menu){
		
		MenuItem deleteMenuItem = new MenuItem(menu, SWT.CASCADE);
		deleteMenuItem.setText("&Delete");
		Menu menuDelete = new Menu(deleteMenuItem);
		deleteMenuItem.setMenu(menuDelete);
		deleteMenuItem.setID(2050000);
		
		/*
		 * Delete All menu item
		 */
		MenuItem menuItemDeleteAll = new MenuItem(menuDelete, SWT.CASCADE);
		menuItemDeleteAll.setText("All");
		menuItemDeleteAll.setID(2050100);
		
		/*
		 * Hide In Sheet menu item
		 */
		MenuItem menuItemDeleteInSheet = new MenuItem(menuDelete, SWT.CASCADE);
		menuItemDeleteInSheet.setText("In Sheet");
		menuItemDeleteInSheet.setID(2050200);
		
		return deleteMenuItem;
	}
	
	private MenuItem addExportMenu(Menu menu){
		
		MenuItem exportMenuItem = new MenuItem(menu, SWT.CASCADE);
		exportMenuItem.setText("&Export as");
		Menu menuExport = new Menu(exportMenuItem);
		exportMenuItem.setMenu(menuExport);
		exportMenuItem.setID(2060000);
		
		/*
		 * Export As CSV menu item
		 */
		MenuItem menuItemExportCSV = new MenuItem(menuExport, SWT.CASCADE);
		menuItemExportCSV.setText("CSV");
		menuItemExportCSV.setID(2060100);
		
		/*
		 * Export As Workbook menu item
		 */
		MenuItem menuItemExcelWorkbook = new MenuItem(menuExport, SWT.CASCADE);
		menuItemExcelWorkbook.setText("Workbook");
		menuItemExcelWorkbook.setID(2060200);
		
		return exportMenuItem;
	}
	
	/**
	 * @return the menuItems
	 */
	public MenuItem[] getMenuItems() {
		return menuItems;
	}
}
