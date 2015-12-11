package de.tudresden.annotator.main;

import org.eclipse.swt.SWT;
import org.eclipse.swt.events.SelectionAdapter;
import org.eclipse.swt.events.SelectionEvent;
import org.eclipse.swt.ole.win32.OleAutomation;
import org.eclipse.swt.widgets.Menu;
import org.eclipse.swt.widgets.MenuItem;
import org.eclipse.swt.widgets.MessageBox;
import org.eclipse.swt.widgets.Shell;

import de.tudresden.annotator.annotations.AnnotationClass;
import de.tudresden.annotator.annotations.ClassGenerator;
import de.tudresden.annotator.oleutils.AnnotationUtils;

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
		
		MenuItem menuFileClose = new MenuItem(menuFile, SWT.CASCADE);
		menuFileClose.setText("Close");
		menuFileClose.addSelectionListener(new SelectionAdapter() {
			@Override
			public void widgetSelected(SelectionEvent e) {
				MainWindow.getInstance().disposeControlSite();
			}
		});		
		//menuFileClose.setEnabled(false);
		
		MenuItem menuFileSave = new MenuItem(menuFile, SWT.CASCADE);
		menuFileSave.setText("Save");
		menuFileSave.addSelectionListener(new SelectionAdapter() {
			@Override
			public void widgetSelected(SelectionEvent e) {
				OleAutomation workbookAutomation = MainWindow.getInstance().getEmbeddedWorkbook();
				String directoryPath = MainWindow.getInstance().getDirectoryPath();
				String fileName = MainWindow.getInstance().getFileName();				
				AnnotationUtils.exportAnnotationsAsCSV(workbookAutomation, directoryPath, fileName);
			}
		});
		//menuFileSave.setEnabled(false);
		
		
		MenuItem menuFileExit = new MenuItem(menuFile, SWT.CASCADE);
		menuFileExit.setText("Exit");
		menuFileExit.addSelectionListener(new SelectionAdapter() {
			@Override
			public void widgetSelected(SelectionEvent e) {		
				MessageBox msgbox = MainWindow.getInstance().createMessageBox(SWT.ICON_QUESTION | SWT.YES | SWT.NO);
				msgbox.setMessage("Are you sure you want to exit the program?");
				
				if(msgbox.open()==SWT.YES){					
					MainWindow.getInstance().disposeControlSite();
					MainWindow.getInstance().disposeShell();
				}
			}
		});
		
		return fileMenu;
	}
	
	private MenuItem addEditMenu(Menu menuBar){
		
		MenuItem editMenu = new MenuItem(menuBar, SWT.CASCADE);
		editMenu.setText("&Edit");
		Menu menuEdit = new Menu(editMenu);
		editMenu.setMenu(menuEdit);
				
		addAnnotateMenu(menuEdit);
		
		MenuItem menuEditAnnotations = new MenuItem(menuEdit, SWT.CASCADE);
		menuEditAnnotations.setText("Annotations");
		menuEditAnnotations.addSelectionListener(new SelectionAdapter() {
			@Override
			public void widgetSelected(SelectionEvent e) {
				//TODO: Create and open window that displays the current annotations for the loaded file.
				//The user shall be able to edit the existing annotations
			}
		});
		
		MenuItem menuEditShowFormulas = new MenuItem(menuEdit, SWT.CASCADE);
		menuEditShowFormulas.setText("Show Formulas");
		menuEditShowFormulas.addSelectionListener(new SelectionAdapter() {
			@Override
			public void widgetSelected(SelectionEvent e) {
				//TODO: Implement as cascade menu having options such as Clear->All , Clear->Specific Annotation
			}
		});
		
		MenuItem menuEditClear = new MenuItem(menuEdit, SWT.CASCADE);
		menuEditClear.setText("Clear ");
		menuEditClear.addSelectionListener(new SelectionAdapter() {
			@Override
			public void widgetSelected(SelectionEvent e) {
				
				OleAutomation workbookAutomation = MainWindow.getInstance().getEmbeddedWorkbook();
				String sheetName = MainWindow.getInstance().getActiveWorksheetName();
				
				AnnotationUtils.clearShapeAnnotationsFromActiveSheet(workbookAutomation, sheetName);
				//AnnotationUtils.clearAnnotationDataForActiveSheet(workbookAutomation, sheetName);
				//TODO: Implement as cascade menu having options such as Clear->All , Clear->Specific Annotation
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

	
	private MenuItem addAnnotateMenu(Menu menuBar){
		
		MenuItem annotateMenu = new MenuItem(menuBar, SWT.CASCADE);
		annotateMenu.setText("&Mark As");
		Menu menuAnnotate = new Menu(annotateMenu);
		annotateMenu.setMenu(menuAnnotate);
		
		AnnotationClass[] classes = ClassGenerator.createAnnotationClasses();
		
		for (AnnotationClass annotationClass : classes) {
			
			MenuItem menuAnnotateTable = new MenuItem(menuAnnotate, SWT.CASCADE);
			menuAnnotateTable.setText(annotationClass.getLabel());
			menuAnnotateTable.addSelectionListener(new SelectionAdapter() {
				@Override
				public void widgetSelected(SelectionEvent e) {
					 
					 OleAutomation workbookAutomation = MainWindow.getInstance().getEmbeddedWorkbook();
					 String sheetName = MainWindow.getInstance().getActiveWorksheetName();
					 int sheetIndex = MainWindow.getInstance().getActiveWorksheetIndex();
					 String[] currentSelection = MainWindow.getInstance().getCurrentSelection();

					 AnnotationUtils.callAnnotationMethod(workbookAutomation, sheetName, sheetIndex, currentSelection, annotationClass);
				}
			});		
		}
		
		return annotateMenu;
	}

	private MenuItem addPreferencesMenu(Menu menuBar){
		MenuItem preferencesMenu = new MenuItem(menuBar, SWT.CASCADE);
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

	public MenuItem[] getMenuItems() {
		return menuItems;
	}
	
	private long getRGBColorAsLong(int red, int green, int blue){	
		return blue * 65536 + green * 256 + red;
	}
	
}
