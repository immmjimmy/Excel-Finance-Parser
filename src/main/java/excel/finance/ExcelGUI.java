package excel.finance;

import java.awt.EventQueue;

import javax.swing.JFrame;
import javax.swing.JPanel;
import java.awt.BorderLayout;
import javax.swing.JButton;
import javax.swing.JFileChooser;

import java.awt.event.ActionListener;
import java.io.File;
import java.io.IOException;
import java.awt.event.ActionEvent;
import javax.swing.JTextField;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import java.util.List;

public class ExcelGUI {

    private JFrame frame;
    private JTextField textField;
    private List<List<List<String>>> parsedData;
    private String filePath;

    /**
     * Launch the application.
     */
    public static void main(String[] args) {
	EventQueue.invokeLater(new Runnable() {
	    public void run() {
		try {
		    ExcelGUI window = new ExcelGUI();
		    window.frame.setVisible(true);
		} catch (Exception e) {
		    e.printStackTrace();
		}
	    }
	});
    }

    /**
     * Create the application.
     */
    public ExcelGUI() {
	initialize();
    }

    /**
     * Initialize the contents of the frame.
     */
    private void initialize() {
	frame = new JFrame();
	frame.setBounds(100, 100, 432, 62);
	frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);

	JPanel panel = new JPanel();
	frame.getContentPane().add(panel, BorderLayout.CENTER);

	JButton openFileBtn = new JButton("Open File");
	openFileBtn.addActionListener(new ActionListener() {
	    public void actionPerformed(ActionEvent e) {
		JFileChooser fileChooser = new JFileChooser();
		if(fileChooser.showOpenDialog(null) == JFileChooser.APPROVE_OPTION) {
		    File file = fileChooser.getSelectedFile();
		    filePath = file.getAbsolutePath();
		    try {
			String extension = "";
			int dotIndex = filePath.lastIndexOf('.');
			int slashIndex = Math.max(filePath.lastIndexOf('/'), filePath.lastIndexOf('\\'));
			
			if(dotIndex > slashIndex) {
			    extension = filePath.substring(dotIndex + 1);
			}
			if(extension.equals("xls") || extension.equals("xlsx")) {
			    parsedData = ExcelParser.readFile(filePath);
			}
		    } catch(InvalidFormatException | IOException ex) {
			ex.printStackTrace();
		    }
		    textField.setText(filePath);
		}
	    }
	});
	panel.add(openFileBtn);

	textField = new JTextField();
	panel.add(textField);
	textField.setColumns(10);

	JButton createFileBtn = new JButton("Create modified data");
	createFileBtn.addActionListener(new ActionListener() {
	    public void actionPerformed(ActionEvent e) {
		if(parsedData != null) {
		    try {
			ExcelParser.createFile(filePath, parsedData);
			textField.setText("Success!");
		    } catch (InvalidFormatException | IOException ex) {
			ex.printStackTrace();
		    }
		} else {
		    textField.setText("Failed!");
		}
	    }
	});
	panel.add(createFileBtn);
    }

}
