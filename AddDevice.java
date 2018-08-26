
/*
    TechCheck -- Simple inventory management tool using a barcode reader to check in and out devices.
    Copyright (C) 2018  MD Showman

    This program is free software: you can redistribute it and/or modify
    it under the terms of the GNU General Public License as published by
    the Free Software Foundation, either version 3 of the License, or
    (at your option) any later version.

    This program is distributed in the hope that it will be useful,
    but WITHOUT ANY WARRANTY; without even the implied warranty of
    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
    GNU General Public License for more details.

    You should have received a copy of the GNU General Public License
    along with this program.  If not, see <https://www.gnu.org/licenses/>. 
 */

import java.awt.BorderLayout;
import java.awt.Dimension;
import java.awt.FlowLayout;

import javax.swing.JButton;
import javax.swing.JDialog;
import javax.swing.JPanel;
import javax.swing.border.EmptyBorder;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.swing.JLabel;
import javax.swing.JOptionPane;
import javax.swing.JTextField;
import javax.swing.JComboBox;
import javax.swing.DefaultComboBoxModel;
import java.awt.event.ActionListener;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.awt.event.ActionEvent;

public class AddDevice extends JDialog {

	private final JPanel barcodePanel = new JPanel();
	private JTextField txtBarcode;
	private JComboBox<String> comboBox;

	private static XSSFWorkbook wb;
	private static XSSFSheet sh;
	private static FileInputStream fis;
	private static FileOutputStream fos;

	/**
	 * Create the dialog.
	 * 
	 * @throws IOException
	 * @throws InvalidFormatException
	 * @throws EncryptedDocumentException
	 */
	public AddDevice() throws EncryptedDocumentException, InvalidFormatException, IOException {
		// call method to create workbook (update)
		connectWorkbook();

		setResizable(false);
		setDefaultCloseOperation(DISPOSE_ON_CLOSE);
		setBounds(100, 100, 450, 137);

		setLocationRelativeTo(null);

		getContentPane().setLayout(new BorderLayout());
		barcodePanel.setLayout(new FlowLayout());
		barcodePanel.setBorder(new EmptyBorder(5, 5, 5, 5));
		getContentPane().add(barcodePanel, BorderLayout.NORTH);

		JLabel lblBarcode = new JLabel("Scan Barcode:");
		barcodePanel.add(lblBarcode);

		txtBarcode = new JTextField();
		barcodePanel.add(txtBarcode);
		txtBarcode.setColumns(27);

		JPanel deviceTypePanel = new JPanel();
		getContentPane().add(deviceTypePanel, BorderLayout.CENTER);

		JLabel lblDeviceType = new JLabel("Device Type:");
		deviceTypePanel.add(lblDeviceType);

		comboBox = new JComboBox<String>();
		comboBox.setModel(new DefaultComboBoxModel(new String[] {"HP Chromebook", "Lenovo Chromebook", "iPad Mini", "Windows Laptop"}));
		comboBox.setSelectedIndex(-1);
		comboBox.setPreferredSize(new Dimension(350, 20));
		deviceTypePanel.add(comboBox);

		JPanel buttonPane = new JPanel();
		buttonPane.setLayout(new FlowLayout(FlowLayout.RIGHT));
		getContentPane().add(buttonPane, BorderLayout.SOUTH);

		JButton btnAdd = new JButton("Add Device");
		btnAdd.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				try {
					addRecord();
				} catch (EncryptedDocumentException e1) {
					// TODO Auto-generated catch block
					e1.printStackTrace();
				} catch (InvalidFormatException e1) {
					// TODO Auto-generated catch block
					e1.printStackTrace();
				}
				resetFields();
			}
		});
		btnAdd.setActionCommand("");
		buttonPane.add(btnAdd);
		getRootPane().setDefaultButton(btnAdd);

		JButton cancelButton = new JButton("Cancel");
		cancelButton.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				try {
					wb.close();
				} catch (IOException e1) {
					// TODO Auto-generated catch block
					e1.printStackTrace();
				}
				dispose();
			}
		});
		cancelButton.setActionCommand("Cancel");
		buttonPane.add(cancelButton);

	}

	public void addRecord() throws EncryptedDocumentException, InvalidFormatException {
		if (alreadyCreated(txtBarcode.getText()) == 1) {

			JOptionPane.showMessageDialog(this, "This device is already in the inventory.", "Add Device",
					JOptionPane.OK_OPTION);
			txtBarcode.requestFocus();
		} else if (txtBarcode.getText() != "" && comboBox.getSelectedIndex() > -1) {
			try {
				sh.createRow(sh.getLastRowNum() + 1);

				for (int i = 0; i < 4; i++) {
					sh.getRow(sh.getLastRowNum()).createCell(i);
				}

				sh.getRow(sh.getLastRowNum()).getCell(0).setCellValue(Integer.parseInt(txtBarcode.getText()));
				sh.getRow(sh.getLastRowNum()).getCell(1).setCellValue(comboBox.getSelectedItem().toString());
				sh.getRow(sh.getLastRowNum()).getCell(2).setCellValue("Available");
				sh.getRow(sh.getLastRowNum()).getCell(3).setCellValue("-");

				fis.close();
				fos = new FileOutputStream("./inventory.xlsx");
				wb.write(fos);
				fos.flush();
				fos.close();

				wb.close();

				connectWorkbook();

				JOptionPane.showMessageDialog(this, "Device added to inventory.", "Add Device", JOptionPane.OK_OPTION);

			} catch (FileNotFoundException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			} catch (IOException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
		} else {
			JOptionPane.showMessageDialog(this, "Please enter a barcode and choose a type", "Add Device",
					JOptionPane.OK_OPTION);
			txtBarcode.requestFocus();
		}
	}

	public static void connectWorkbook() throws EncryptedDocumentException, InvalidFormatException, IOException {
		fis = new FileInputStream("./inventory.xlsx");
		new WorkbookFactory();
		wb = (XSSFWorkbook) WorkbookFactory.create(fis);

		DataFormat fmt = wb.createDataFormat();
		CellStyle textStyle = wb.createCellStyle();
		textStyle.setDataFormat(fmt.getFormat("@"));

		sh = wb.getSheet("Inventory");
		sh.setDefaultColumnStyle(0, textStyle);
	}

	public void resetFields() {
		txtBarcode.setText("");
		comboBox.setSelectedIndex(-1);
		txtBarcode.requestFocus();
	}

	public int alreadyCreated(String id) {
		for (int i = 1; i <= sh.getLastRowNum(); i++) {
			int temp = (int) sh.getRow(i).getCell(0).getNumericCellValue();

			if (String.valueOf(temp).equals(id))
				return 1;
		}

		return 0;
	}

}
