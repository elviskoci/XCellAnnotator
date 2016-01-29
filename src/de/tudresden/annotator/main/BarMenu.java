package de.tudresden.annotator.main;

import java.util.ArrayList;
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
		menuFileSave.addSelectionListener(GUIListeners.createFileSaveSelectionListener());
		
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
		menuFileClose.addSelectionListener(GUIListeners.createFileCloseSelectionListener());
		
		/*
		 *  Exit Application menu item 
		 */
		MenuItem menuFileExit = new MenuItem(menuFile, SWT.CASCADE);
		menuFileExit.setID(1070000);
		menuFileExit.setText("Exit \tCtrl+Q");
		menuFileExit.setAccelerator(SWT.MOD1 + 'Q');
		menuFileExit.addSelectionListener(GUIListeners.createFileExitSelectionListener());
		
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
		 * Show Annotations menu item  
		 */
		MenuItem menuItemShow = addShowMenu(menuAnnotations);
		menuItemShow.setEnabled(false);
		
		
		/*
		 * Delete annotations menu item  
		 */
		MenuItem menuItemDelete = addDeleteMenu(menuAnnotations);
		menuItemDelete.setEnabled(false);
		
		/*
		 * Show Formulas menu item  
		 */
		MenuItem menuItemShowFormulas = new MenuItem(menuAnnotations, SWT.CHECK);
		menuItemShowFormulas.setID(2060000);
		menuItemShowFormulas.setText("Show Formulas\tCtrl+F2");
		menuItemShowFormulas.setEnabled(false);
		menuItemShowFormulas.addSelectionListener(GUIListeners.createShowFormulasSelectionListener());
		menuItemShowFormulas.setAccelerator(SWT.MOD1 + SWT.F2);
		
		/*
		 * Show Annotations menu item  
		 */
		MenuItem menuItemUndo = new MenuItem(menuAnnotations, SWT.CASCADE);
		menuItemUndo.setID(2080000);
		menuItemUndo.setText("Undo Annotation\tCtrl+Z");
		menuItemUndo.setEnabled(false);
		menuItemUndo.addSelectionListener(GUIListeners.createUndoLastAnnotationSelectionListener());
		menuItemUndo.setAccelerator(SWT.MOD1+'Z');
		
		MenuItem menuItemRedo = new MenuItem(menuAnnotations, SWT.CASCADE);
		menuItemRedo.setID(2090000);
		menuItemRedo.setText("Redo Annotation\tCtrl+Y");
		menuItemRedo.setEnabled(false);
		menuItemRedo.addSelectionListener(GUIListeners.createRedoLastAnnotationSelectionListener());
		menuItemRedo.setAccelerator(SWT.MOD1+'Y');
		
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
		menuViewFolderExplorer.setEnabled(false);
		
		/*
		 * View Annotation Management Panel menu item  
		 */
		MenuItem menuViewAnnotationsPanel = new MenuItem(menuView, SWT.CASCADE);
		menuViewAnnotationsPanel.setText("Annotations Panel");
		menuViewAnnotationsPanel.setID(3020000);
		menuViewAnnotationsPanel.setEnabled(false);
		
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
		menuPreferencesFormatting.setEnabled(false);
		
		/*
		 * Annotation Classes Preferences menu item 
		 */
		MenuItem menuPreferencesAnnotationClasses = new MenuItem(menuPreferences, SWT.CASCADE);
		menuPreferencesAnnotationClasses.setText("Annotation Classes");
		menuPreferencesAnnotationClasses.setID(4020000);
		menuPreferencesAnnotationClasses.setEnabled(false);
		
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
		usedChars.add('W'); // sheet completed
		usedChars.add('E'); // sheet not applicable
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
		MenuItem menuItemNotApplicable = new MenuItem(menuAnnotateWorksheet, SWT.CHECK);
		menuItemNotApplicable.setID(2020100);
		// menuItemNotApplicable.setText("Not Applicable \tCtrl+E");
		menuItemNotApplicable.setText("Not Applicable \tE");
		menuItemNotApplicable.addSelectionListener(GUIListeners.createSheetNotApplicableSelectionListener());
		menuItemNotApplicable.setAccelerator('E');
		// menuItemNotApplicable.setAccelerator(SWT.MOD1+'E');
		
		/*
		 * Worksheet Completed menu item
		 */
		MenuItem menuItemCompleted = new MenuItem(menuAnnotateWorksheet, SWT.CHECK);
		menuItemCompleted.setID(2020200);
		// menuItemCompleted.setText("Completed \tCtrl+W");
		menuItemCompleted.setText("Completed \tW");
		menuItemCompleted.addSelectionListener(GUIListeners.createSheetCompletedSelectionListener());
		menuItemCompleted.setAccelerator('W');
		// menuItemCompleted.setAccelerator(SWT.MOD1+'W');
		
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
		MenuItem menuItemNotApplicable = new MenuItem(menuAnnotateWorkbook, SWT.CHECK);
		menuItemNotApplicable.setID(2030100);
		menuItemNotApplicable.setText("Not Applicable");
		// menuItemNotApplicable.setText("Not Applicable \tCtrl+Shift+E");
		menuItemNotApplicable.addSelectionListener(GUIListeners.createFileNotApplicableSelectionListener());
		// menuItemNotApplicable.setAccelerator(SWT.MOD1+SWT.MOD2+'E');
		
		/*
		 * Workbook Completed menu item 
		 */
		MenuItem menuItemCompleted = new MenuItem(menuAnnotateWorkbook, SWT.CHECK);
		menuItemCompleted.setID(2030200);
		menuItemCompleted.setText("Completed");
		// menuItemCompleted.setText("Completed \tCtrl+Shift+W");
		menuItemCompleted.addSelectionListener(GUIListeners.createFileCompletedSelectionListener());
		// menuItemCompleted.setAccelerator(SWT.MOD1+SWT.MOD2+'W');
		return annotateWorkbookMenuItem;
	}
	
	private MenuItem addHideMenu(Menu menu){
		
		MenuItem hideMenuItem = new MenuItem(menu, SWT.CASCADE);
		hideMenuItem.setID(2040000);
		hideMenuItem.setText("&Hide");
		Menu menuHide = new Menu(hideMenuItem);
		hideMenuItem.setMenu(menuHide);
		
		/*
		 * Hide All menu item
		 */
		MenuItem menuItemHideAll = new MenuItem(menuHide, SWT.CASCADE);
		menuItemHideAll.setID(2040100);
		menuItemHideAll.setText("All");
		menuItemHideAll.addSelectionListener(GUIListeners.createHideAllAnnotationsSelectionListener());
		
		/*
		 * Hide In Sheet menu item
		 */
		MenuItem menuItemHideInSheet = new MenuItem(menuHide, SWT.CASCADE);
		menuItemHideInSheet.setID(2040200);
		menuItemHideInSheet.setText("In Sheet");
		menuItemHideInSheet.addSelectionListener(GUIListeners.createHideInSheetAnnotationsSelectionListener());
	
		return hideMenuItem;
	}
	
	private MenuItem addShowMenu(Menu menu){
		
		MenuItem showMenuItem = new MenuItem(menu, SWT.CASCADE);
		showMenuItem.setID(2070000);
		showMenuItem.setText("&Show");
		Menu menuShow = new Menu(showMenuItem);
		showMenuItem.setMenu(menuShow);
		
		/*
		 * Show All menu item
		 */
		MenuItem menuItemShowAll = new MenuItem(menuShow, SWT.CASCADE);
		menuItemShowAll.setID(2070100);
		menuItemShowAll.setText("All");
		menuItemShowAll.addSelectionListener(GUIListeners.createShowAllAnnotationsSelectionListener());
		
		/*
		 * Show In Sheet menu item
		 */
		MenuItem menuItemHideInSheet = new MenuItem(menuShow, SWT.CASCADE);
		menuItemHideInSheet.setID(2070700);
		menuItemHideInSheet.setText("In Sheet");
		menuItemHideInSheet.addSelectionListener(GUIListeners.createShowInSheetAnnotationsSelectionListener());
	
		return showMenuItem;
	}

	private MenuItem addDeleteMenu(Menu menu){
		
		MenuItem deleteMenuItem = new MenuItem(menu, SWT.CASCADE);
		deleteMenuItem.setID(2050000);
		deleteMenuItem.setText("&Delete");
		Menu menuDelete = new Menu(deleteMenuItem);
		deleteMenuItem.setMenu(menuDelete);
		
		/*
		 * Delete All menu item
		 */
		MenuItem menuItemDeleteAll = new MenuItem(menuDelete, SWT.CASCADE);
		menuItemDeleteAll.setID(2050100);
		menuItemDeleteAll.setText("All");
		menuItemDeleteAll.addSelectionListener(GUIListeners.createDeleteAllAnnotationsSelectionListener());
		
		/*
		 * Delete In Sheet menu item
		 */
		MenuItem menuItemDeleteInSheet = new MenuItem(menuDelete, SWT.CASCADE);
		menuItemDeleteInSheet.setID(2050200);
		menuItemDeleteInSheet.setText("In Sheet");
		menuItemDeleteInSheet.addSelectionListener(GUIListeners.createDeleteAnnotationsInSheetSelectionListener());
		
		/*
		 * Delete In Sheet menu item
		 */
		MenuItem menuItemDeleteInRange = new MenuItem(menuDelete, SWT.CASCADE);
		menuItemDeleteInRange.setID(2050300);
		menuItemDeleteInRange.setText("In Range");
		menuItemDeleteInRange.addSelectionListener(GUIListeners.createDeleteAnnotationsInRangeSelectionListener());
		
		return deleteMenuItem;
	}
	
	private MenuItem addExportMenu(Menu menu){
		
		MenuItem exportMenuItem = new MenuItem(menu, SWT.CASCADE);
		exportMenuItem.setID(2060000);
		exportMenuItem.setText("&Export as");
		Menu menuExport = new Menu(exportMenuItem);
		exportMenuItem.setMenu(menuExport);
		
		/*
		 * Export As CSV menu item
		 */
		MenuItem menuItemExportCSV = new MenuItem(menuExport, SWT.CASCADE);
		menuItemExportCSV.setID(2060100);
		menuItemExportCSV.setText("CSV");
		menuItemExportCSV.addSelectionListener(GUIListeners.createExportAsCSVSelectionListener());
		
		/*
		 * Export As Workbook menu item
		 */
		MenuItem menuItemExcelWorkbook = new MenuItem(menuExport, SWT.CASCADE);
		menuItemExcelWorkbook.setID(2060200);
		menuItemExcelWorkbook.setText("Workbook");
		menuItemExcelWorkbook.addSelectionListener(GUIListeners.createExportAsWorkbookSelectionListener());
		
		return exportMenuItem;
	}
	
	/**
	 * @return the menuItems
	 */
	public MenuItem[] getMenuItems() {
		return menuItems;
	}
}
