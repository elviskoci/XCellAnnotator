package de.tudresden.annotator.main;

import java.util.Iterator;

import org.eclipse.swt.SWT;
import org.eclipse.swt.events.SelectionAdapter;
import org.eclipse.swt.events.SelectionEvent;
import org.eclipse.swt.ole.win32.OleAutomation;
import org.eclipse.swt.widgets.Menu;
import org.eclipse.swt.widgets.MenuItem;
import org.eclipse.swt.widgets.MessageBox;
import org.eclipse.swt.widgets.Shell;

import de.tudresden.annotator.annotations.AnnotationClass;
import de.tudresden.annotator.annotations.utils.AnnotationDataSheet;
import de.tudresden.annotator.annotations.utils.AnnotationHandler;
import de.tudresden.annotator.annotations.utils.AnnotationResult;
import de.tudresden.annotator.annotations.utils.ClassGenerator;
import de.tudresden.annotator.annotations.utils.ValidationResult;
import de.tudresden.annotator.oleutils.ApplicationUtils;
import de.tudresden.annotator.oleutils.FileUtils;
import de.tudresden.annotator.oleutils.WorkbookUtils;

public class BarMenu {
	
	private MenuItem[] menuItems = new MenuItem[4];
	
	public BarMenu(Shell oleShell){
		
		final Shell shell = oleShell;
		
		Menu menuBar = shell.getMenuBar();
		if (menuBar == null) {
			menuBar = new Menu(shell, SWT.BAR);
			shell.setMenuBar(menuBar);
		}
	
		menuItems[0] = addFileMenu(menuBar);
		menuItems[1] = addEditMenu(menuBar);
		menuItems[2] = addViewMenu(menuBar);
		menuItems[3] = addPreferencesMenu(menuBar);
	}
	
	
	private MenuItem addFileMenu(Menu menuBar){
		
		MenuItem fileMenu = new MenuItem(menuBar, SWT.CASCADE);
		fileMenu.setText("&File");
		Menu menuFile = new Menu(fileMenu);
		fileMenu.setMenu(menuFile);
		
		MenuItem menuFileOpen = new MenuItem(menuFile, SWT.CASCADE);
		menuFileOpen.setText("Open...");
		menuFileOpen.addSelectionListener(new SelectionAdapter() {
			@Override
			public void widgetSelected(SelectionEvent e) {
				MainWindow.getInstance().fileOpen();
				OleAutomation workbookAutomation = MainWindow.getInstance().getEmbeddedWorkbook();
				boolean isDataInMemory = AnnotationDataSheet.readAnnotationData(workbookAutomation);
				if(isDataInMemory){
					System.out.println(AnnotationHandler.getWorkbookAnnotation().getWorksheetAnnotations());
					AnnotationHandler.setVisilityForShapeAnnotations(workbookAutomation, true);
					AnnotationDataSheet.setVisibility(workbookAutomation, true);
				}else{
					int style = SWT.ICON_WARNING;
	        		MessageBox messageBox = MainWindow.getInstance().createMessageBox(style);
	 	            messageBox.setMessage("Could not read the annotation data. "
	 	            		+ "Either there are no annotations or the annotation data are not in the expect format.");
	 	            messageBox.open();
				}
			}
		});
		
		MenuItem menuFileOpenPrevious = new MenuItem(menuFile, SWT.CASCADE);
		menuFileOpenPrevious.setText("Open Prev");
		menuFileOpenPrevious.addSelectionListener(new SelectionAdapter() {
			@Override
			public void widgetSelected(SelectionEvent e) {
				//TODO: Open the next Excel file from the current folder.   
			}
		});
		menuFileOpenPrevious.setEnabled(false);
		
		MenuItem menuFileOpenNext = new MenuItem(menuFile, SWT.CASCADE);
		menuFileOpenNext.setText("Open Next");
		menuFileOpenNext.addSelectionListener(new SelectionAdapter() {
			@Override
			public void widgetSelected(SelectionEvent e) {
				//TODO: Open the next Excel file from the current folder.   
			}
		});
		menuFileOpenNext.setEnabled(false);
				
		MenuItem menuFileSave = new MenuItem(menuFile, SWT.CASCADE);
		menuFileSave.setText("Save");
		menuFileSave.addSelectionListener(new SelectionAdapter() {
			@Override
			public void widgetSelected(SelectionEvent e) {
				
				OleAutomation embeddedWorkbook = MainWindow.getInstance().getEmbeddedWorkbook();
				String fileName = MainWindow.getInstance().getFileName();
				String directory = MainWindow.getInstance().getDirectoryPath();
				String filePath = directory+"\\"+fileName;
				
				ApplicationUtils.setDisplayAlerts(MainWindow.getInstance().getExcelApplication(), "False");		
				boolean result = FileUtils.saveProgress(embeddedWorkbook, filePath);
				if(result){
            		FileUtils.markFileAsAnnotated(directory, fileName, 1);
				}else{
					System.out.println("The file was not saved!");
				}
				
				ApplicationUtils.setDisplayAlerts(MainWindow.getInstance().getExcelApplication(), "True");	
			}
		});
		//menuFileSave.setEnabled(false);
		
		addExportMenu(menuFile);
		
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
		//menuFileClose.setEnabled(false);
		
		MenuItem menuFileExit = new MenuItem(menuFile, SWT.CASCADE);
		menuFileExit.setText("Exit");
		menuFileExit.addSelectionListener(new SelectionAdapter() {
			@Override
			public void widgetSelected(SelectionEvent e) {	
				
				if( MainWindow.getInstance().getControlSite()!=null && 
						MainWindow.getInstance().getControlSite().isDirty() &&
						 	MainWindow.getInstance().getEmbeddedWorkbook()!=null){
					
					
					// System.out.println(AnnotationHandler.getWorkbookAnnotation().toString());
				    // System.out.println( AnnotationHandler.getWorkbookAnnotation().getWorksheetAnnotations().size());
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
	 	            		System.err.println("Could not save progress!");
	 	            	}else{
	 	            		String directory = MainWindow.getInstance().getDirectoryPath();
	 	            		String fileName = MainWindow.getInstance().getFileName();
	 	            		FileUtils.markFileAsAnnotated(directory, fileName, 1);
	 	            		
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
		
		return fileMenu;
	}
	
	private MenuItem addEditMenu(Menu menuBar){
		
		MenuItem editMenu = new MenuItem(menuBar, SWT.CASCADE);
		editMenu.setText("&Annotations");
		Menu menuEdit = new Menu(editMenu);
		editMenu.setMenu(menuEdit);
				
		addAnnotateMenu(menuEdit);
			
		
		MenuItem menuEditAnnotations = new MenuItem(menuEdit, SWT.CASCADE);
		menuEditAnnotations.setText("Edit");
		menuEditAnnotations.addSelectionListener(new SelectionAdapter() {
			@Override
			public void widgetSelected(SelectionEvent e) {
				//TODO: Create and open window that displays the current annotations for the loaded file.
				//The user shall be able to edit the existing annotations
			}
		});
		
		
		addHideMenu(menuEdit);
		
		addDeleteMenu(menuEdit);
		
		MenuItem menuEditShowFormulas = new MenuItem(menuEdit, SWT.CASCADE);
		menuEditShowFormulas.setText("Show Formulas");
		menuEditShowFormulas.addSelectionListener(new SelectionAdapter() {
			@Override
			public void widgetSelected(SelectionEvent e) {
				//TODO: Implement as cascade menu having options such as Clear->All , Clear->Specific Annotation
			}
		});
		
		
		MenuItem menuEditShowAnnotations = new MenuItem(menuEdit, SWT.CASCADE);
		menuEditShowAnnotations.setText("Show Annotations");
		menuEditShowAnnotations.addSelectionListener(new SelectionAdapter() {
			@Override
			public void widgetSelected(SelectionEvent e) {
				OleAutomation embeddedWorkbook =  MainWindow.getInstance().getEmbeddedWorkbook();
				AnnotationHandler.setVisilityForShapeAnnotations(embeddedWorkbook, true);
				AnnotationDataSheet.setVisibility(embeddedWorkbook, true);
			}
		});
		
		return editMenu;
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

	
	private MenuItem addAnnotateMenu(Menu menu){
		
		MenuItem markAsMenuItem = new MenuItem(menu, SWT.CASCADE);
		markAsMenuItem.setText("&Mark As");
		Menu menuMarkAs = new Menu(markAsMenuItem);
		markAsMenuItem.setMenu(menuMarkAs);
		
		Iterator<AnnotationClass> itr = ClassGenerator.getAnnotationClasses().values().iterator();
	
		while(itr.hasNext()){
			AnnotationClass annotationClass = (AnnotationClass) itr.next();
			MenuItem menuAnnotateTable = new MenuItem(menuMarkAs, SWT.CASCADE);
			menuAnnotateTable.setText(annotationClass.getLabel());
			menuAnnotateTable.addSelectionListener(new SelectionAdapter() {
				@Override
				public void widgetSelected(SelectionEvent e) {
					 
					 OleAutomation workbookAutomation = MainWindow.getInstance().getEmbeddedWorkbook();
					 String sheetName = MainWindow.getInstance().getActiveWorksheetName();
					 int sheetIndex = MainWindow.getInstance().getActiveWorksheetIndex();
					 String[] currentSelection = MainWindow.getInstance().getCurrentSelection();

					 AnnotationResult  result=  AnnotationHandler.annotate(workbookAutomation, sheetName, sheetIndex, currentSelection, annotationClass);
					 
					 if(result.getValidationResult()!=ValidationResult.OK){
		        		MessageBox messageBox = MainWindow.getInstance().createMessageBox(SWT.ICON_ERROR);
		 	            messageBox.setMessage(result.getMessage());
		 	            messageBox.open();
					 }
				}
			});		
		}	
		return markAsMenuItem;
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
		menuItemClearInSheet.setText("Sheet");
		menuItemClearInSheet.addSelectionListener(new SelectionAdapter() {
			@Override
			public void widgetSelected(SelectionEvent e) {
				
				OleAutomation applicationAutomation = MainWindow.getInstance().getExcelApplication();
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
				AnnotationDataSheet.deleteAllAnnotationData(workbookAutomation);
			}
		});	
		
		
		MenuItem menuItemDeleteInSheet = new MenuItem(menuDelete, SWT.CASCADE);
		menuItemDeleteInSheet.setText("Sheet");
		menuItemDeleteInSheet.addSelectionListener(new SelectionAdapter() {
			@Override
			public void widgetSelected(SelectionEvent e){
				//OleAutomation applicationAutomation = MainWindow.getInstance().getExcelApplication();
				OleAutomation workbookAutomation = MainWindow.getInstance().getEmbeddedWorkbook();
				String sheetName = MainWindow.getInstance().getActiveWorksheetName();
				
				AnnotationHandler.deleteShapeAnnotationsFromWorksheet(workbookAutomation, sheetName);
				AnnotationDataSheet.deleteAnnotationDataForWorksheet(workbookAutomation, sheetName);
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
				boolean isSuccess = AnnotationDataSheet.exportAnnotationsAsCSV(workbookAutomation, directoryPath, fileName);
				
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
	
	
	private long getRGBColorAsLong(int red, int green, int blue){	
		return blue * 65536 + green * 256 + red;
	}
	
}
