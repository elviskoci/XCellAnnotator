/**
 * 
 */
package de.tudresden.annotator.main;

import java.io.File;

import org.eclipse.swt.SWT;
import org.eclipse.swt.SWTError;
import org.eclipse.swt.layout.FillLayout;
import org.eclipse.swt.ole.win32.OLE;
import org.eclipse.swt.ole.win32.OleControlSite;
import org.eclipse.swt.ole.win32.OleFrame;
import org.eclipse.swt.widgets.Display;
import org.eclipse.swt.widgets.FileDialog;
import org.eclipse.swt.widgets.MessageBox;
import org.eclipse.swt.widgets.Shell;

/**
 * @author Elvis
 *
 */
public class MainWindow {
	
	// Sink GUID 
	private static final String IID_AppEvents = "{00024413-0000-0000-C000-000000000046}";
	// Event ID
	private static final int SheetSelectionChange   = 0x00000616;
	
	
	private static final Display display = new Display();
	private static final Shell shell = new Shell(display);
	private OleFrame oleFrame;
	private OleControlSite controlSite;
	
	private String currentSelection[];
	private String activeWorksheetName;
	private long activeWorksheetIndex;
	
	
	private static MainWindow instance = null;
	protected MainWindow() {
      // Exists only to defeat instantiation.
    }
  
	public static MainWindow getInstance() {
		if(instance == null) {
			instance = new MainWindow();
		}
		return instance;  
	}
	
	/**
	 * Create the window that will serve as the main Graphical User Interface (GUI)  
	 * @param shell
	 */
	private void buildGUIWindow(Shell shell){
		shell.setText("Annotator");
	    shell.setLayout(new FillLayout());
	    shell.setSize(1200, 600);
		
	    oleFrame = new OleFrame(shell, SWT.NONE);
	    
	    BarMenu  oleFrameMenuBar = new BarMenu(oleFrame.getShell());
	    oleFrame.setFileMenus(oleFrameMenuBar.getMenuItems());
	}
	
	
	/**
	 * Open an excel file for annotation
	 */
	 public void fileOpen(){
		
		// Select the excel file
		Shell shell = oleFrame.getShell();
		FileDialog dialog = new FileDialog(shell, SWT.OPEN);
		String fileName = dialog.open();
		
		// if no file was selected, return
		if (fileName == null) return;
		
		// dispose OleControlSite if it is not null. 
		disposeControlSite();
				
	    if (controlSite == null) {
			int index = fileName.lastIndexOf('.');
			if (index != -1) {
				String fileExtension = fileName.substring(index + 1);
				if (fileExtension.equalsIgnoreCase("xls") || fileExtension.equalsIgnoreCase("xlsx")) {	
					
					// create a new control site to open the file with Excel
					try {		    	
				        File excelFile = new File(fileName);
				        controlSite = new OleControlSite(oleFrame, SWT.NONE, excelFile);
				        controlSite.doVerb(OLE.OLEIVERB_INPLACEACTIVATE);				    
				    } catch (SWTError e) {
				        e.printStackTrace();
				        System.out.println("Unable to open ActiveX Control");
				        return;
				    }	    	  
				   
				}else{
					MessageBox msgbox = new MessageBox(shell,SWT.ICON_ERROR);
					msgbox.setMessage("The selected file format is not recognized: ."+fileExtension);
					msgbox.open();
				}
			}
	    }
	}
	
	
	/**
	 * Dispose control site 
	 */
	public void disposeControlSite() {
		if (controlSite != null){
			controlSite.dispose();
		}
		controlSite = null;
	}
	
	/**
	 * Dispose shell
	 */
	public void disposeShell() {
		if (shell != null){
			shell.dispose();
		}
	}
	
	/**
	 * Get OleFrame
	 * @return
	 */
	public OleFrame getOleFrame() {
		return oleFrame;
	}
		
	/**
	 * Set OleFrame
	 * @param oleFrame
	 */
	public void setOleFrame(OleFrame oleFrame) {
		this.oleFrame = oleFrame;
	}
	
	/**
	 * Get OleControlSite
	 * @return
	 */
	public OleControlSite getControlSite() {
		return controlSite;
	}
	
	/**
	 * Set OleControlSite
	 * @param controlSite
	 */
	public void setControlSite(OleControlSite controlSite) {
		this.controlSite = controlSite;
	}
	
	/**
	 * Create message box using the "main" window (this class) shell 
	 * @param style 
	 * @return
	 */
	public MessageBox createMessageBox(int style){
		return new MessageBox(shell,style);
	}
	
	/**
	 * @param args
	 */
	public static void main(String[] args) {

		MainWindow gui = MainWindow.getInstance(); 
	    gui.buildGUIWindow(shell);

  		shell.open();
  		
  	    while (!shell.isDisposed ()) {
  	        if (!display.readAndDispatch ()) display.sleep();
  	    }
  	    
	    display.dispose();
	}
			
}
