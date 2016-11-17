package br.com.degasperi.access2excel.ui;

import org.apache.commons.io.FilenameUtils;
import org.eclipse.jface.action.MenuManager;
import org.eclipse.jface.action.StatusLineManager;
import org.eclipse.jface.action.ToolBarManager;
import org.eclipse.jface.dialogs.MessageDialog;
import org.eclipse.jface.window.ApplicationWindow;
import org.eclipse.swt.SWT;
import org.eclipse.swt.events.SelectionAdapter;
import org.eclipse.swt.events.SelectionEvent;
import org.eclipse.swt.graphics.Point;
import org.eclipse.swt.widgets.Button;
import org.eclipse.swt.widgets.Combo;
import org.eclipse.swt.widgets.Composite;
import org.eclipse.swt.widgets.Control;
import org.eclipse.swt.widgets.Display;
import org.eclipse.swt.widgets.FileDialog;
import org.eclipse.swt.widgets.Label;
import org.eclipse.swt.widgets.Shell;
import org.eclipse.swt.widgets.Text;

import br.com.degasperi.access2excel.classes.Access2Excel;
import br.com.degasperi.access2excel.classes.Access2Excel.OutputFormat;
import br.com.degasperi.access2excel.classes.Access2Excel.StepWriter;

public class MainWindow extends ApplicationWindow {
	private Text fileName;

	/**
	 * Create the application window.
	 */
	public MainWindow() {
		super(null);
		setShellStyle(SWT.CLOSE | SWT.MIN | SWT.APPLICATION_MODAL);
		createActions();
		// addToolBar(SWT.FLAT | SWT.WRAP);
		// addMenuBar();
		// addStatusLine();
	}

	/**
	 * Create contents of the application window.
	 * 
	 * @param parent
	 */
	@Override
	protected Control createContents(Composite parent) {
		Composite container = new Composite(parent, SWT.NONE);

		fileName = new Text(container, SWT.BORDER);
		fileName.setEditable(false);
		fileName.setBounds(10, 31, 256, 21);

		Button bOpen = new Button(container, SWT.NONE);
		bOpen.addSelectionListener(new SelectionAdapter() {
			@Override
			public void widgetSelected(SelectionEvent e) {
				FileDialog fd = new FileDialog(container.getShell(), SWT.OPEN);
				fd.setFilterExtensions(new String[] { "*.mdb", "*.accdb" });
				fd.setFilterNames(new String[] { "MDB Files (*.mdb)", "ACCDB Files (*.accdb)" });
				fd.setText("Choose a file...");
				String filename = fd.open();
				if (filename != null)
					fileName.setText(filename);
			}
		});
		bOpen.setBounds(266, 30, 23, 23);
		bOpen.setText("...");

		Label lblChooseAMdb = new Label(container, SWT.NONE);
		lblChooseAMdb.setBounds(10, 10, 125, 15);
		lblChooseAMdb.setText("Choose a file:");

		Combo combo = new Combo(container, SWT.READ_ONLY);
		combo.setItems(new String[] { "XLS", "XLSX" });
		combo.setBounds(10, 79, 278, 23);
		combo.select(0);

		Label lblNewLabel = new Label(container, SWT.NONE);
		lblNewLabel.setBounds(10, 58, 91, 15);
		lblNewLabel.setText("Output format:");

		Label lStatus = new Label(container, SWT.NONE);
		lStatus.setEnabled(false);
		lStatus.setBounds(10, 169, 278, 15);

		Button bConvert = new Button(container, SWT.NONE);
		bConvert.addSelectionListener(new SelectionAdapter() {
			@Override
			public void widgetSelected(SelectionEvent e) {
				try (Access2Excel converter = new Access2Excel()) {
					converter.setOutputFormat(OutputFormat.valueOf(combo.getText()));
					converter.setWriter(new MainStepWriter(lStatus));
					converter.open(fileName.getText());
					String outputFileName = FilenameUtils.removeExtension(fileName.getText()) + "."
							+ combo.getText().toLowerCase();
					converter.convert(outputFileName);
					MessageDialog.openInformation(Display.getCurrent().getActiveShell(), "Conversion successfull",
							"File save in " + outputFileName + ".");
				} catch (Exception e1) {
					MessageDialog.openError(Display.getCurrent().getActiveShell(), "Conversion error", e1.getMessage());
				}
			}
		});
		bConvert.setBounds(10, 112, 278, 51);
		bConvert.setText("Convert");

		return container;
	}

	/**
	 * Create the actions.
	 */
	private void createActions() {
		// Create the actions
	}

	/**
	 * Create the menu manager.
	 * 
	 * @return the menu manager
	 */
	@Override
	protected MenuManager createMenuManager() {
		MenuManager menuManager = new MenuManager("menu");
		return menuManager;
	}

	/**
	 * Create the toolbar manager.
	 * 
	 * @return the toolbar manager
	 */
	@Override
	protected ToolBarManager createToolBarManager(int style) {
		ToolBarManager toolBarManager = new ToolBarManager(style);
		return toolBarManager;
	}

	/**
	 * Create the status line manager.
	 * 
	 * @return the status line manager
	 */
	@Override
	protected StatusLineManager createStatusLineManager() {
		StatusLineManager statusLineManager = new StatusLineManager();
		return statusLineManager;
	}

	/**
	 * Launch the application.
	 * 
	 * @param args
	 */
	public static void main(String args[]) {
		try {
			MainWindow window = new MainWindow();
			window.setBlockOnOpen(true);
			window.open();
			Display.getCurrent().dispose();
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	/**
	 * Configure the shell.
	 * 
	 * @param newShell
	 */
	@Override
	protected void configureShell(Shell newShell) {
		super.configureShell(newShell);
		newShell.setText("Access2Excel");
	}

	/**
	 * Return the initial size of the window.
	 */
	@Override
	protected Point getInitialSize() {
		return new Point(304, 220);
	}

	public static class MainStepWriter implements StepWriter {
		private Label label;

		public MainStepWriter(Label label) {
			this.label = label;
		}

		@Override
		public void write(String text) {
			StepWriter.super.write(text);
			label.setText(text);
		}

		@Override
		public void error(String text) {
			throw new RuntimeException(text);
		}
	}
}
