package de.tudresden.annotator.main;

import java.util.ArrayList;
import java.util.Arrays;
import java.util.Iterator;
import java.util.LinkedHashMap;

import org.eclipse.swt.SWT;
import org.eclipse.swt.events.SelectionAdapter;
import org.eclipse.swt.events.SelectionEvent;
import org.eclipse.swt.ole.win32.OleAutomation;
import org.eclipse.swt.widgets.Menu;
import org.eclipse.swt.widgets.MenuItem;
import org.eclipse.swt.widgets.MessageBox;
import org.eclipse.swt.widgets.Shell;

import de.tudresden.annotator.annotations.AnnotationClass;
import de.tudresden.annotator.annotations.WorkbookAnnotation;
import de.tudresden.annotator.annotations.WorksheetAnnotation;
import de.tudresden.annotator.annotations.utils.RangeAnnotationsSheet;
import de.tudresden.annotator.annotations.utils.AnnotationHandler;
import de.tudresden.annotator.annotations.utils.AnnotationResult;
import de.tudresden.annotator.annotations.utils.ClassGenerator;
import de.tudresden.annotator.annotations.utils.ValidationResult;
import de.tudresden.annotator.oleutils.ApplicationUtils;
import de.tudresden.annotator.oleutils.WorkbookUtils;
import de.tudresden.annotator.oleutils.WorksheetUtils;

public class BarMenuOLD {
	
	private MenuItem[] menuItems = new MenuItem[4];
	
	public BarMenuOLD(Shell oleShell){
		
		final Shell shell = oleShell;
		
		Menu menuBar = shell.getMenuBar();
		if (menuBar == null) {
			menuBar = new Menu(shell, SWT.BAR);
			shell.setMenuBar(menuBar);
		}
	
		menuItems[0] = addFileMenu(menuBar);
		menuItems[1] = addAnnotationsMenu(menuBar);
		menuItems[2] = addViewMenu(menuBar);
		menuItems[3] = addPreferencesMenu(menuBar);
	}
	
	
	private MenuItem addFileMenu(Menu menuBar){
		
		MenuItem fileMenu = new MenuItem(menuBar, SWT.CASCADE);
		fileMenu.setText("&File");
		Menu menuFile = new Menu(fileMenu);
		fileMenu.setMenu(menuFile);
		
		/*
		 *  open file
		 */
		MenuItem menuFileOpen = new MenuItem(menuFile, SWT.CASCADE);
		menuFileOpen.setText("Open... \tCtrl+O");
		menuFileOpen.addSelectionListener(new SelectionAdapter() {
			@Override
			public void widgetSelected(SelectionEvent e) {
				
				// open the files selection window
				FileUtils.fileOpen();
				
				// get the OleAutomation for the embedded workbook
				OleAutomation workbookAutomation = MainWindow.getInstance().getEmbeddedWorkbook();
				if (workbookAutomation == null) {
					return;
				}
				
				boolean isProtected = WorkbookUtils.protectWorkbook(workbookAutomation, true, false);		
				if(!isProtected){
					int style = SWT.ERROR;
					MessageBox message = MainWindow.getInstance().createMessageBox(style);
					message.setMessage("ERROR: Could not protect the workbook. Operation failed!");
					message.open();
					WorkbookUtils.closeEmbeddedWorkbook(workbookAutomation, false);
					MainWindow.getInstance().disposeControlSite();
					return;
				}
				
				
				OleAutomation  annotationDataSheet = 
						WorkbookUtils.getWorksheetAutomationByName(workbookAutomation, RangeAnnotationsSheet.getName()); 
				
				// protect the annotation data sheet if it is not protected already
				if(annotationDataSheet!=null){
					isProtected = WorksheetUtils.protectWorksheet(annotationDataSheet);
					if(!isProtected){
						int style = SWT.ERROR;
						MessageBox message = MainWindow.getInstance().createMessageBox(style);
						message.setMessage("ERROR: Could not protect annotation data sheet! ");
						message.open();
						WorkbookUtils.closeEmbeddedWorkbook(workbookAutomation, false);
						MainWindow.getInstance().disposeControlSite();
						return;
					}
				}
				
				// show the annotation data sheet
				// TODO: Check why it does not make visible
				if(annotationDataSheet!=null)
					WorksheetUtils.setWorksheetVisibility(annotationDataSheet, true);
				// read the annotation data and recreate inmemory structure
				RangeAnnotationsSheet.readRangeAnnotations(workbookAutomation);
				
				// re-draw all the annotation in memory structure 
				AnnotationHandler.drawAllAnnotations(workbookAutomation);
								
				// enable menu items (sub-menus) that are relevant  
				MenuItem fileMenu = menuItems[0]; // File menu 
				MenuItem[] fileSubmenus = fileMenu.getMenu().getItems();
				for (MenuItem menuItem : fileSubmenus) {
					if(!(menuItem.getText().compareTo("Open Prev")==0 || menuItem.getText().compareTo("Open Next")==0)){
						menuItem.setEnabled(true);
					}	
				}
					
				MenuItem annotationsMenu = menuItems[1]; // Annotations menu 
				MenuItem[] annotationsSubmenus = annotationsMenu.getMenu().getItems();
				for (MenuItem menuItem : annotationsSubmenus) {
					menuItem.setEnabled(true);
				}
			}
		});
		menuFileOpen.setAccelerator(SWT.MOD1+'O');
		
		/*
		 *  open prev file
		 */
		MenuItem menuFileOpenPrevious = new MenuItem(menuFile, SWT.CASCADE);
		menuFileOpenPrevious.setText("Open Prev");
		menuFileOpenPrevious.addSelectionListener(new SelectionAdapter() {
			@Override
			public void widgetSelected(SelectionEvent e) {
				System.out.println("Prev File");
			}
		});
		menuFileOpenPrevious.setEnabled(false);
		
		/*
		 *  open next file
		 */
		MenuItem menuFileOpenNext = new MenuItem(menuFile, SWT.CASCADE);
		menuFileOpenNext.setText("Open Next");
		menuFileOpenNext.addSelectionListener(new SelectionAdapter() {
			@Override
			public void widgetSelected(SelectionEvent e) {
				System.out.println("Next File");
			}
		});
		menuFileOpenNext.setEnabled(false);
				
		/*
		 *  save file
		 */
		MenuItem menuFileSave = new MenuItem(menuFile, SWT.CASCADE);
		menuFileSave.setText("Save \tCtrl+S");
		menuFileSave.addSelectionListener(new SelectionAdapter() {
			@Override
			public void widgetSelected(SelectionEvent e) {
				
				OleAutomation embeddedWorkbook = MainWindow.getInstance().getEmbeddedWorkbook();
				String fileName = MainWindow.getInstance().getFileName();
				String directory = MainWindow.getInstance().getDirectoryPath();
				String filePath = directory+"\\"+fileName;
				
				OleAutomation applicationAutomation = WorkbookUtils.getApplicationAutomation(embeddedWorkbook);
				
				ApplicationUtils.setDisplayAlerts(applicationAutomation, "False");		
				boolean result = FileUtils.saveProgress(embeddedWorkbook, filePath);
				if(result){
            		//FileUtils.markFileAsAnnotated(directory, fileName, 1);
            		int style = SWT.ICON_INFORMATION;
					MessageBox message = MainWindow.getInstance().createMessageBox(style);
					message.setMessage("The file was successfully saved.");
					message.open();
				}else{
					int style = SWT.ICON_ERROR;
					MessageBox message = MainWindow.getInstance().createMessageBox(style);
					message.setMessage("ERROR: The file could not be saved!");
					message.open();
				}
				
				ApplicationUtils.setDisplayAlerts(applicationAutomation, "True");	
			}
		});
		menuFileSave.setEnabled(false);
		menuFileSave.setAccelerator(SWT.MOD1 + 'S');
		
		/*
		 *  export file
		 */
		MenuItem menuFileExport = addExportMenu(menuFile);
		menuFileExport.setEnabled(false);
		
		/*
		 *  close file
		 */
		MenuItem menuFileClose = new MenuItem(menuFile, SWT.CASCADE);
		menuFileClose.setText("Close");
		menuFileClose.addSelectionListener(new SelectionAdapter() {
			@Override
			public void widgetSelected(SelectionEvent e) {
				OleAutomation embeddedWorkbook  = MainWindow.getInstance().getEmbeddedWorkbook();
				WorkbookUtils.closeEmbeddedWorkbook(embeddedWorkbook, false);
				MainWindow.getInstance().disposeControlSite();
			}
		});		
		menuFileClose.setEnabled(false);
		
		
		/*
		 *  exit application
		 */
		MenuItem menuFileExit = new MenuItem(menuFile, SWT.CASCADE);
		menuFileExit.setText("Exit \tCtrl+Q");
		menuFileExit.addSelectionListener(new SelectionAdapter() {
			@Override
			public void widgetSelected(SelectionEvent e) {	
				
				if( !MainWindow.getInstance().isControlSiteNull() && 
						MainWindow.getInstance().isControlSiteDirty() &&
						 	MainWindow.getInstance().getEmbeddedWorkbook()!=null){
					
					
					System.out.println( AnnotationHandler.getWorkbookAnnotation().getWorksheetAnnotations());
					
	        		int style = SWT.YES | SWT.NO | SWT.CANCEL | SWT.ICON_WARNING ;
	        		MessageBox messageBox = MainWindow.getInstance().createMessageBox(style);
	 	            messageBox.setMessage("Want to save your changes?");
	 	            
	 	            int response = messageBox.open();	 	 	            
	 	            if( response== SWT.YES){	
	 	            	OleAutomation embeddedWorkbook = MainWindow.getInstance().getEmbeddedWorkbook();
	 	            	String filePath =  MainWindow.getInstance().getDirectoryPath()+"\\"+MainWindow.getInstance().getFileName();
	 	            	boolean isSaved = FileUtils.saveProgress(embeddedWorkbook, filePath);
	 	            	
	 	            	if(!isSaved){
	 	            		int messageStyle = SWT.ICON_ERROR;
	 						MessageBox message = MainWindow.getInstance().createMessageBox(messageStyle);
	 						message.setMessage("ERROR: The file could not be saved!");
	 						message.open();
	 	            	}else{
	 	            		//String directory = MainWindow.getInstance().getDirectoryPath();
	 	            		//String fileName = MainWindow.getInstance().getFileName();
	 	            		//FileUtils.markFileAsAnnotated(directory, fileName, 1);
	 	            		
		 	            	WorkbookUtils.closeEmbeddedWorkbook(embeddedWorkbook, false);
		 	            	MainWindow.getInstance().disposeControlSite();
		 	            	MainWindow.getInstance().disposeShell();
	 	            	}
	 	            } 
	 	            
	 	            if(response == SWT.NO){
	 	            	MainWindow.getInstance().disposeControlSite();
	 	            	MainWindow.getInstance().disposeShell();
	 	            } 
	        	}else{
	        		int style = SWT.YES | SWT.NO | SWT.ICON_QUESTION;
	        		MessageBox messageBox = MainWindow.getInstance().createMessageBox(style);
	 	            messageBox.setMessage("Do you want to close the application?");
	 	            
	 	            int response = messageBox.open();
	 	            if( response== SWT.YES){
	 	        	    MainWindow.getInstance().disposeControlSite();
	 	            	MainWindow.getInstance().disposeShell();
	 	            }
	        	}
			}
		});
		menuFileExit.setAccelerator(SWT.MOD1 + 'Q');
		
		return fileMenu;
	}
	
	private MenuItem addAnnotationsMenu(Menu menuBar){
		
		MenuItem annotationsMenu = new MenuItem(menuBar, SWT.CASCADE);
		annotationsMenu.setText("&Annotations");
		Menu menuAnnotations = new Menu(annotationsMenu);
		annotationsMenu.setMenu(menuAnnotations);
				
		MenuItem menuItemRange = addAnnotateRangeMenu(menuAnnotations);
		menuItemRange.setEnabled(false);
		
		MenuItem menuItemSheet = addAnnotateWorksheetMenu(menuAnnotations);	
		menuItemSheet.setEnabled(false);
		
		MenuItem menuItemBook = addAnnotateWorkbookMenu(menuAnnotations);
		menuItemBook.setEnabled(false);
				
		MenuItem menuItemHide = addHideMenu(menuAnnotations);
		menuItemHide.setEnabled(false);
		
		MenuItem menuItemDelete = addDeleteMenu(menuAnnotations);
		menuItemDelete.setEnabled(false);
		
		MenuItem menuItemShowFormulas = new MenuItem(menuAnnotations, SWT.CASCADE);
		menuItemShowFormulas.setText("Show Formulas");
		menuItemShowFormulas.addSelectionListener(new SelectionAdapter() {
			@Override
			public void widgetSelected(SelectionEvent e) {
				//TODO: Implement as cascade menu having options such as Clear->All , Clear->Specific Annotation
			}
		});
		menuItemShowFormulas.setEnabled(false);
		
		MenuItem menuItemShowAnnotations = new MenuItem(menuAnnotations, SWT.CASCADE);
		menuItemShowAnnotations.setText("Show Annotations");
		menuItemShowAnnotations.addSelectionListener(new SelectionAdapter() {
			@Override
			public void widgetSelected(SelectionEvent e) {
				OleAutomation embeddedWorkbook =  MainWindow.getInstance().getEmbeddedWorkbook();
				AnnotationHandler.setVisilityForShapeAnnotations(embeddedWorkbook, true);
				RangeAnnotationsSheet.setVisibility(embeddedWorkbook, true);
			}
		});
		menuItemShowAnnotations.setEnabled(false);
		
		return annotationsMenu;
	}
	
	
	private MenuItem addViewMenu(Menu menuBar) {
		MenuItem viewMenu = new MenuItem(menuBar, SWT.CASCADE);
		viewMenu.setText("&View");
		Menu menuView = new Menu(viewMenu);
		viewMenu.setMenu(menuView);
		
		MenuItem menuViewFolderExplorer = new MenuItem(menuView, SWT.CASCADE);
		menuViewFolderExplorer.setText("Folder Explorer");
		menuViewFolderExplorer.addSelectionListener(new SelectionAdapter() {
			@Override
			public void widgetSelected(SelectionEvent e) {
				//TODO: Hide/Show Folder Explorer Panel  
			}
		});
		
		MenuItem menuViewAnnotationsPanel = new MenuItem(menuView, SWT.CASCADE);
		menuViewAnnotationsPanel.setText("Annotations Panel");
		menuViewAnnotationsPanel.addSelectionListener(new SelectionAdapter() {
			@Override
			public void widgetSelected(SelectionEvent e) {
				//TODO: Hide/Show Annotations Panel  
			}
		});
		
		return viewMenu;
	}

	
	private MenuItem addAnnotateRangeMenu(Menu menu){
		
		MenuItem annotateRangeMenuItem = new MenuItem(menu, SWT.CASCADE);
		annotateRangeMenuItem.setText("&Range as");
		Menu menuAnnotateRange = new Menu(annotateRangeMenuItem);
		annotateRangeMenuItem.setMenu(menuAnnotateRange);
		
		LinkedHashMap<String, AnnotationClass> map =  ClassGenerator.getAnnotationClasses();
		Iterator<String> keys = map.keySet().iterator();
		
		// using the leftmost characters to make it easier for the user to simultaneously handle the mouse and keyboard
		ArrayList<Character> shortcutChars  = new ArrayList<Character>();
		shortcutChars.addAll(Arrays.asList(new Character[]{'A', 'S', 'D', 'X', 'Z', 'C', 'Q', 'W', 'E', 'F'}));
		ArrayList<Character> usedChars = new ArrayList<Character>(); 
		
		while(keys.hasNext()){
			String label = keys.next();
			if(shortcutChars.contains(label.charAt(0))){
				AnnotationClass ac = map.get(label);
				ac.setShortcut(SWT.MOD1 | SWT.MOD2 + label.charAt(0));
				usedChars.add(label.charAt(0));
			}
		}
		
		for (int i = 0; i < usedChars.size(); i++) {
			shortcutChars.remove(usedChars.get(i));
		}
		
		Iterator<AnnotationClass> values = map.values().iterator();
		int i = 0;
		while(values.hasNext()){
			AnnotationClass annotationClass = (AnnotationClass) values.next();
			MenuItem menuAnnotationClass = new MenuItem(menuAnnotateRange, SWT.CASCADE);
			menuAnnotationClass.addSelectionListener(new SelectionAdapter() {
				@Override
				public void widgetSelected(SelectionEvent e) {
					 
					 OleAutomation workbookAutomation = MainWindow.getInstance().getEmbeddedWorkbook();
					 
					 String sheetName = MainWindow.getInstance().getActiveWorksheetName();
					 int sheetIndex = MainWindow.getInstance().getActiveWorksheetIndex();
					 String[] currentSelection = MainWindow.getInstance().getCurrentSelection();

					 AnnotationResult  result=  
							 AnnotationHandler.annotate(workbookAutomation, sheetName, sheetIndex,   
					 		 currentSelection, annotationClass);				 
					 
					 if(result.getValidationResult()!=ValidationResult.OK){
		        		MessageBox messageBox = MainWindow.getInstance().createMessageBox(SWT.ICON_ERROR);
		 	            messageBox.setMessage(result.getMessage());
		 	            messageBox.open();
					 }
					 				 
					 MainWindow.getInstance().setFocusToControlSite();				 
				}
			});		
			int shortcut = annotationClass.getShortcut();
			if(shortcut < 0){
				shortcut = SWT.MOD1 | SWT.MOD2 + shortcutChars.get(i);
				annotationClass.setShortcut(shortcut);
			}
			menuAnnotationClass.setAccelerator(shortcut);
			char ch = (char) (shortcut - SWT.MOD1 | SWT.MOD2);
			menuAnnotationClass.setText(annotationClass.getLabel()+"\t Ctrl+Shift+"+ch);
			
			i++;
		}	
		return annotateRangeMenuItem;
	}
	
	private MenuItem addAnnotateWorksheetMenu(Menu menu){
		
		MenuItem annotateWorksheetMenuItem = new MenuItem(menu, SWT.CASCADE);
		annotateWorksheetMenuItem.setText("&Sheet as");
		Menu menuAnnotateWorksheet = new Menu(annotateWorksheetMenuItem);
		annotateWorksheetMenuItem.setMenu(menuAnnotateWorksheet);
		
		MenuItem menuItemNotApplicable = new MenuItem(menuAnnotateWorksheet, SWT.CASCADE);
		menuItemNotApplicable.setText("Not Applicable");
		menuItemNotApplicable.addSelectionListener(new SelectionAdapter() {
			@Override
			public void widgetSelected(SelectionEvent e) {
				
				String sheetName = MainWindow.getInstance().getActiveWorksheetName();
				WorkbookAnnotation workbookAnnotation = AnnotationHandler.getWorkbookAnnotation();
				WorksheetAnnotation sheetAnnotation = workbookAnnotation.getWorksheetAnnotations().get(sheetName);
							
				String strMessage = ""; 
				if(sheetAnnotation.isNotApplicable()){
					sheetAnnotation.setNotApplicable(false);
					strMessage = "The \"Not Applicable\" status was removed for sheet \""+sheetName+"\". The sheet can now be annotated";
					
					MenuItem annotationsMenu = menuItems[1]; // Annotations menu 
					MenuItem[] annotationsSubmenus = annotationsMenu.getMenu().getItems();				
					for (MenuItem menuItem : annotationsSubmenus) {
							menuItem.setEnabled(true);
					}
					annotateWorksheetMenuItem.getMenu().getItem(1).setEnabled(true); // enable "Completed" menu item. 
						
				}else{
					sheetAnnotation.setNotApplicable(true);
					
					MenuItem annotationsMenu = menuItems[1]; // Annotations menu 
					MenuItem[] annotationsSubmenus = annotationsMenu.getMenu().getItems();				
					for (MenuItem menuItem : annotationsSubmenus) {
						if(menuItem.getText().compareTo("&Range as")==0 || menuItem.getText().compareTo("&Delete")==0 ||
							menuItem.getText().compareTo("&Hide")==0 || menuItem.getText().compareTo("Show Annotations")==0){
							
							menuItem.setEnabled(false);
						}else{
							menuItem.setEnabled(true);
						}
					}
					annotateWorksheetMenuItem.getMenu().getItem(1).setEnabled(false); // disable "Completed" menu item. 
					
					strMessage = "The worksheet \""+sheetName+"\" was marked as \"Not Applicable\"";
				}
				
				int style = SWT.ICON_INFORMATION;
				MessageBox mb = MainWindow.getInstance().createMessageBox(style);
				mb.setMessage(strMessage); 
				mb.open();
			}
		});	
		
		MenuItem menuItemCompleted = new MenuItem(menuAnnotateWorksheet, SWT.CASCADE);
		menuItemCompleted.setText("Completed");
		menuItemCompleted.addSelectionListener(new SelectionAdapter() {
			@Override
			public void widgetSelected(SelectionEvent e) {
				String sheetName = MainWindow.getInstance().getActiveWorksheetName();
				WorkbookAnnotation workbookAnnotation = AnnotationHandler.getWorkbookAnnotation();
				WorksheetAnnotation sheetAnnotation = workbookAnnotation.getWorksheetAnnotations().get(sheetName);
				
				String strMessage = ""; 
				if(sheetAnnotation.isCompleted()){
					sheetAnnotation.setCompleted(false);
					
					MenuItem annotationsMenu = menuItems[1]; // Annotations menu 
					MenuItem[] annotationsSubmenus = annotationsMenu.getMenu().getItems();
					for (MenuItem menuItem : annotationsSubmenus) {
							menuItem.setEnabled(true);
					}
					annotateWorksheetMenuItem.getMenu().getItem(0).setEnabled(true); // enable "NotApplicable" menu item.					
					strMessage = "The status for the sheet \""+sheetName+"\" was changed back to \"Not Completed\"";
					
				}else{
					
					sheetAnnotation.setCompleted(true);
					
					MenuItem annotationsMenu = menuItems[1]; // Annotations menu 
					MenuItem[] annotationsSubmenus = annotationsMenu.getMenu().getItems();
					for (MenuItem menuItem : annotationsSubmenus) {
						if(menuItem.getText().compareTo("&Range as")==0 || menuItem.getText().compareTo("&Delete")==0){
							menuItem.setEnabled(false);
						}else{
							menuItem.setEnabled(true);
						}
					}
					annotateWorksheetMenuItem.getMenu().getItem(0).setEnabled(false); // disable "NotApplicable" menu item.
					
					strMessage = "The sheet \""+sheetName+"\" was marked completed";
				}
				
				int style = SWT.ICON_INFORMATION;
				MessageBox mb = MainWindow.getInstance().createMessageBox(style);
				mb.setMessage(strMessage); 
				mb.open();
			}
		});	
		
		return annotateWorksheetMenuItem;
	}
	
	
	private MenuItem addAnnotateWorkbookMenu(Menu menu){
		
		MenuItem annotateWorkbookMenuItem = new MenuItem(menu, SWT.CASCADE);
		annotateWorkbookMenuItem.setText("&File as");
		Menu menuAnnotateWorkbook = new Menu(annotateWorkbookMenuItem);
		annotateWorkbookMenuItem.setMenu(menuAnnotateWorkbook);
		
		MenuItem menuItemIrrelevant = new MenuItem(menuAnnotateWorkbook, SWT.CASCADE);
		menuItemIrrelevant.setText("Not Applicable");
		menuItemIrrelevant.addSelectionListener(new SelectionAdapter() {
			@Override
			public void widgetSelected(SelectionEvent e) {
				WorkbookAnnotation workbookAnnotation = AnnotationHandler.getWorkbookAnnotation();
				String fileName = MainWindow.getInstance().getFileName();
				
				String strMessage = ""; 
				if(workbookAnnotation.isNotApplicable()){
					workbookAnnotation.setNotApplicable(false);
					
					MenuItem annotationsMenu = menuItems[1]; // Annotations menu 
					MenuItem[] annotationsSubmenus = annotationsMenu.getMenu().getItems();
					for (MenuItem menuItem : annotationsSubmenus) {
							menuItem.setEnabled(true);
					}
					annotateWorkbookMenuItem.getMenu().getItem(1).setEnabled(true); // enable "Completed" menu item.
					
					strMessage = "The \"Not Applicable\" status was removed for file \""+fileName+"\". "
							+ "You can now proceed with the annotations.";
				}else{
					workbookAnnotation.setNotApplicable(true);
					
					MenuItem annotationsMenu = menuItems[1]; // Annotations menu 
					MenuItem[] annotationsSubmenus = annotationsMenu.getMenu().getItems();
					for (MenuItem menuItem : annotationsSubmenus) {
						if(menuItem.getText().compareTo("&File as")!=0){
							menuItem.setEnabled(false);
						}else{
							menuItem.setEnabled(true);
						}
					}
					annotateWorkbookMenuItem.getMenu().getItem(1).setEnabled(false); // disable "Completed" menu item.
					
					strMessage = "The file \""+fileName+"\" was marked as \"Not Applicable\"";
				}
				
				int style = SWT.ICON_INFORMATION;
				MessageBox mb = MainWindow.getInstance().createMessageBox(style);
				mb.setMessage(strMessage); 
				mb.open();
				
			}
		});	
		
		MenuItem menuItemCompleted = new MenuItem(menuAnnotateWorkbook, SWT.CASCADE);
		menuItemCompleted.setText("Completed");
		menuItemCompleted.addSelectionListener(new SelectionAdapter() {
			@Override
			public void widgetSelected(SelectionEvent e) {
				
				WorkbookAnnotation workbookAnnotation = AnnotationHandler.getWorkbookAnnotation();
				String fileName = MainWindow.getInstance().getFileName();
				String sheetName = MainWindow.getInstance().getActiveWorksheetName();
				WorksheetAnnotation wa =workbookAnnotation.getWorksheetAnnotations().get(sheetName);
				
				
				String strMessage = ""; 
				if(workbookAnnotation.isCompleted()){
					workbookAnnotation.setCompleted(false);
					
					MenuItem annotationsMenu = menuItems[1]; // Annotations menu 
					MenuItem[] annotationsSubmenus = annotationsMenu.getMenu().getItems();
					for (MenuItem menuItem : annotationsSubmenus) {
						menuItem.setEnabled(true);
					}
					annotateWorkbookMenuItem.getMenu().getItem(0).setEnabled(true); // enable "Not Applicable" menu item.
					
					strMessage = "The status for the file \""+fileName+"\" was changed back to \"Not Completed\"";
				}else{
					workbookAnnotation.setCompleted(true);
					
					MenuItem annotationsMenu = menuItems[1]; // Annotations menu 
					MenuItem[] annotationsSubmenus = annotationsMenu.getMenu().getItems();
					for (MenuItem menuItem : annotationsSubmenus) {
						if(menuItem.getText().compareTo("&Range as")==0 || 
						   menuItem.getText().compareTo("&Sheet as")==0 ||
						   menuItem.getText().compareTo("&Delete")==0	){
								
								menuItem.setEnabled(false);
						}else{
								menuItem.setEnabled(true);
						}
					}
					annotateWorkbookMenuItem.getMenu().getItem(0).setEnabled(false); // disable "Not Applicable" menu item.
					
					strMessage = "The file \""+fileName+"\" was marked as \"Completed\"";
				}
				
				int style = SWT.ICON_INFORMATION;
				MessageBox mb = MainWindow.getInstance().createMessageBox(style);
				mb.setMessage(strMessage); 
				mb.open();
			}
		});	
		
		return annotateWorkbookMenuItem;
	}
	
	private MenuItem addHideMenu(Menu menu){
		
		MenuItem hideMenuItem = new MenuItem(menu, SWT.CASCADE);
		hideMenuItem.setText("&Hide");
		Menu menuHide = new Menu(hideMenuItem);
		hideMenuItem.setMenu(menuHide);
		
		MenuItem menuItemClearAll = new MenuItem(menuHide, SWT.CASCADE);
		menuItemClearAll.setText("All");
		menuItemClearAll.addSelectionListener(new SelectionAdapter() {
			@Override
			public void widgetSelected(SelectionEvent e) {
				OleAutomation workbookAutomation = MainWindow.getInstance().getEmbeddedWorkbook();	
				AnnotationHandler.setVisilityForShapeAnnotations(workbookAutomation, false);
			}
		});	
			
		MenuItem menuItemClearInSheet = new MenuItem(menuHide, SWT.CASCADE);
		menuItemClearInSheet.setText("In Sheet");
		menuItemClearInSheet.addSelectionListener(new SelectionAdapter() {
			@Override
			public void widgetSelected(SelectionEvent e) {

				OleAutomation workbookAutomation = MainWindow.getInstance().getEmbeddedWorkbook();
				String sheetName = MainWindow.getInstance().getActiveWorksheetName();
				AnnotationHandler.setVisibilityForWorksheetShapeAnnotations(workbookAutomation, sheetName, false);
			}
		});	
				
		return hideMenuItem;
	}
	
	
	private MenuItem addDeleteMenu(Menu menu){
		
		MenuItem deleteMenuItem = new MenuItem(menu, SWT.CASCADE);
		deleteMenuItem.setText("&Delete");
		Menu menuDelete = new Menu(deleteMenuItem);
		deleteMenuItem.setMenu(menuDelete);
		
		
		MenuItem menuItemDeleteAll = new MenuItem(menuDelete, SWT.CASCADE);
		menuItemDeleteAll.setText("All");
		menuItemDeleteAll.addSelectionListener(new SelectionAdapter() {
			@Override
			public void widgetSelected(SelectionEvent e) {
				OleAutomation workbookAutomation = MainWindow.getInstance().getEmbeddedWorkbook();	

				AnnotationHandler.deleteAllShapeAnnotations(workbookAutomation);
				
				WorkbookAnnotation workbookAnnotation = AnnotationHandler.getWorkbookAnnotation();
				workbookAnnotation.removeAllAnnotations();
				
				RangeAnnotationsSheet.deleteAllRangeAnnotations(workbookAutomation);
			}
		});	
		
		
		MenuItem menuItemDeleteInSheet = new MenuItem(menuDelete, SWT.CASCADE);
		menuItemDeleteInSheet.setText("In Sheet");
		menuItemDeleteInSheet.addSelectionListener(new SelectionAdapter() {
			@Override
			public void widgetSelected(SelectionEvent e){
				//OleAutomation applicationAutomation = MainWindow.getInstance().getExcelApplication();
				OleAutomation workbookAutomation = MainWindow.getInstance().getEmbeddedWorkbook();
				String sheetName = MainWindow.getInstance().getActiveWorksheetName();
				
				AnnotationHandler.deleteShapeAnnotationsFromWorksheet(workbookAutomation, sheetName);
				
				WorkbookAnnotation workbookAnnotation = AnnotationHandler.getWorkbookAnnotation();
				workbookAnnotation.removeAllAnnotationsFromSheet(sheetName);
				
				RangeAnnotationsSheet.deleteRangeAnnotationsForWorksheet(workbookAutomation, sheetName, false);			
			}
		});	
			
		return deleteMenuItem;
	}

		
	private MenuItem addPreferencesMenu(Menu menu){
		MenuItem preferencesMenu = new MenuItem(menu, SWT.CASCADE);
		preferencesMenu.setText("&Preferences");
		Menu menuPreferences = new Menu(preferencesMenu);
		preferencesMenu.setMenu(menuPreferences);
		
		MenuItem menuPreferencesFormatting = new MenuItem(menuPreferences, SWT.CASCADE);
		menuPreferencesFormatting.setText("Formatting");
		menuPreferencesFormatting.addSelectionListener(new SelectionAdapter() {
			@Override
			public void widgetSelected(SelectionEvent e) {
				//TODO: Create a window that will allow the user to change the color, width, and 
				// other formatting characteristics for each annotation class 
			}
		});
		
		MenuItem menuPreferencesAnnotationClasses = new MenuItem(menuPreferences, SWT.CASCADE);
		menuPreferencesAnnotationClasses.setText("Annotation Classes");
		menuPreferencesAnnotationClasses.addSelectionListener(new SelectionAdapter() {
			@Override
			public void widgetSelected(SelectionEvent e) {
				//TODO: Create, modify, remove annotation classes (Ex: metadata, header) 
			}
		});
		
		return preferencesMenu;
	}
	

	private MenuItem addExportMenu(Menu menu){
		
		MenuItem exportMenuItem = new MenuItem(menu, SWT.CASCADE);
		exportMenuItem.setText("&Export as");
		Menu menuExport = new Menu(exportMenuItem);
		exportMenuItem.setMenu(menuExport);
		
		MenuItem menuItemExportCSV = new MenuItem(menuExport, SWT.CASCADE);
		menuItemExportCSV.setText("CSV");
		menuItemExportCSV.addSelectionListener(new SelectionAdapter() {
			@Override
			public void widgetSelected(SelectionEvent e) {
				OleAutomation workbookAutomation = MainWindow.getInstance().getEmbeddedWorkbook();
				String directoryPath = MainWindow.getInstance().getDirectoryPath();
				String fileName = MainWindow.getInstance().getFileName();				
				boolean isSuccess = RangeAnnotationsSheet.exportRangeAnnotationsAsCSV(workbookAutomation, directoryPath, fileName);
				
				if(isSuccess){
					MessageBox messageBox = MainWindow.getInstance().createMessageBox(SWT.ICON_INFORMATION);
					messageBox.setText("Information");
		            messageBox.setMessage("The annotation data were successfully exported at:\n"+directoryPath);
		            messageBox.open();
				}else{
					MessageBox messageBox = MainWindow.getInstance().createMessageBox(SWT.ICON_ERROR);
					messageBox.setText("Error Message");
		            messageBox.setMessage("An error occured. Could not export the annotation data as csv.");
		            messageBox.open();
				}
			}
		});	
		
		MenuItem menuItemExcelWorkbook = new MenuItem(menuExport, SWT.CASCADE);
		menuItemExcelWorkbook.setText("Workbook");
		menuItemExcelWorkbook.addSelectionListener(new SelectionAdapter() {
			@Override
			public void widgetSelected(SelectionEvent e) {
					MessageBox messageBox = MainWindow.getInstance().createMessageBox(SWT.ICON_INFORMATION);
					messageBox.setText("Information");
		            messageBox.setMessage("This option is not implemented yet");
			}
		});	
		
		
		return exportMenuItem;
	}
	
	
	public MenuItem[] getMenuItems() {
		return menuItems;
	}
}
