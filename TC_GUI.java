
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
import java.awt.EventQueue;

import javax.swing.JFrame;
import javax.swing.JPanel;
import javax.swing.border.EmptyBorder;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import javax.swing.JLabel;
import javax.swing.JOptionPane;
import javax.swing.JTextField;
import javax.swing.JButton;
import java.awt.GridLayout;
import java.awt.event.ActionListener;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.awt.event.ActionEvent;
import javax.swing.JTabbedPane;

public class TC_GUI extends JFrame {
	// GUI
	private JPanel contentPane;
	private JTextField txtBarcode;
	private JTextField txtBorrowedBy;
	private JTextField txtID;
	private JTextField txtType;
	private JTextField txtDate;
	private JButton btnFind, btnCheckOut, btnCheckIn, btnCancel, btnSubmit;

	// non-GUI
	private static Workbook wb;
	private static Sheet sh;
	private static FileInputStream fis;
	private static FileOutputStream fos;

	/**
	 * Launch the application.
	 */
	public static void main(String[] args) throws Exception {
		EventQueue.invokeLater(new Runnable() {
			public void run() {
				try {
					TC_GUI frame = new TC_GUI();
					frame.setVisible(true);

					fis = new FileInputStream("./inventory.xlsx");
					new WorkbookFactory();
					wb = WorkbookFactory.create(fis);
					sh = wb.getSheet("Inventory");

				} catch (Exception e) {
					e.printStackTrace();
				}
			}
		});
	}

	/**
	 * Create the frame.
	 */
	public TC_GUI() {
		setResizable(false);
		setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
		setBounds(100, 100, 450, 300);
		contentPane = new JPanel();
		contentPane.setBorder(new EmptyBorder(5, 5, 5, 5));
		contentPane.setLayout(new BorderLayout(0, 0));
		setContentPane(contentPane);

		setLocationRelativeTo(null);

		// Barcode Panel
		JPanel scanPanel = new JPanel();
		contentPane.add(scanPanel, BorderLayout.NORTH);

		JLabel lblBarcode = new JLabel("Barcode:");
		scanPanel.add(lblBarcode);

		txtBarcode = new JTextField();
		scanPanel.add(txtBarcode);
		txtBarcode.setColumns(10);

		btnFind = new JButton("Search Inventory");
		btnFind.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {

				// Get index for asset
				int foundAsset = getAsset(txtBarcode.getText());

				// If found, populate fields
				if (foundAsset != 0) {
					// System.out.println("adding values");
					txtID.setText(String.valueOf((int) sh.getRow(foundAsset).getCell(0).getNumericCellValue()));
					txtType.setText(sh.getRow(foundAsset).getCell(1).toString());
					txtBorrowedBy.setText(sh.getRow(foundAsset).getCell(2).toString());
					txtDate.setText(sh.getRow(foundAsset).getCell(3).toString());

					// Logic to display CheckOut or CheckIn button
					if (txtBorrowedBy.getText().equalsIgnoreCase("Available")) {
						btnCheckOut.setVisible(true);
						btnCancel.setVisible(true);
					} else {
						btnCheckIn.setVisible(true);
						btnCancel.setVisible(true);
					}
					txtBarcode.setEnabled(false);
					btnFind.setEnabled(false);
				}

				// Reset Barcode panel
				txtBarcode.setText("");
				txtBarcode.requestFocus();
			}
		});
		scanPanel.add(btnFind);

		// Main Panel
		JPanel panel_1 = new JPanel();
		contentPane.add(panel_1, BorderLayout.CENTER);
		panel_1.setLayout(new GridLayout(0, 1, 0, 0));

		// ID Panel
		JPanel idPanel = new JPanel();
		panel_1.add(idPanel);
		idPanel.setLayout(new GridLayout(0, 2, 0, 0));

		JLabel lblAssetIdNumber = new JLabel("Asset ID Number:");
		idPanel.add(lblAssetIdNumber);

		txtID = new JTextField();
		txtID.setEnabled(false);
		txtID.setEditable(false);
		idPanel.add(txtID);
		txtID.setColumns(10);

		// Type Panel
		JPanel typePanel = new JPanel();
		panel_1.add(typePanel);
		typePanel.setLayout(new GridLayout(0, 2, 0, 0));

		JLabel lblType = new JLabel("Type:");
		typePanel.add(lblType);

		txtType = new JTextField();
		txtType.setEnabled(false);
		txtType.setEditable(false);
		typePanel.add(txtType);
		txtType.setColumns(10);

		// BorrowedBy Panel
		JPanel byPanel = new JPanel();
		panel_1.add(byPanel);
		byPanel.setLayout(new GridLayout(0, 2, 0, 0));

		JLabel lblBorrowedBy = new JLabel("Borrowed By:");
		byPanel.add(lblBorrowedBy);

		txtBorrowedBy = new JTextField();
		txtBorrowedBy.setEnabled(false);
		txtBorrowedBy.setEditable(false);
		byPanel.add(txtBorrowedBy);
		txtBorrowedBy.setColumns(10);

		// Date Panel
		JPanel datePanel = new JPanel();
		panel_1.add(datePanel);
		datePanel.setLayout(new GridLayout(0, 2, 0, 0));

		JLabel lblDateBorrowed = new JLabel("Date Borrowed:");
		datePanel.add(lblDateBorrowed);

		txtDate = new JTextField();
		txtDate.setEnabled(false);
		txtDate.setEditable(false);
		datePanel.add(txtDate);
		txtDate.setColumns(10);

		// Buttons panel
		JPanel checkPanel = new JPanel();
		contentPane.add(checkPanel, BorderLayout.SOUTH);

		btnCheckOut = new JButton("Check Out");
		btnCheckOut.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				// Show CheckOut button, populate date field, enable BorrowedBy textfield and
				// gives focus
				btnCheckOut.setVisible(false);
				btnSubmit.setVisible(true);
				String date = new SimpleDateFormat("MM/dd/yyyy").format(new Date());
				txtDate.setText(date);
				txtBorrowedBy.setEnabled(true);
				txtBorrowedBy.requestFocus();
				txtBorrowedBy.setText("");
				txtBorrowedBy.setEditable(true);
			}
		});
		btnCheckOut.setVisible(false);

		btnSubmit = new JButton("Submit");
		btnSubmit.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				// Checks for blank field, if isn't -- write to file, if is -- display warning
				// and set focus to BorrowedBy textfield
				if (!txtBorrowedBy.getText().equals("")) {
					int asset = getAsset(txtID.getText());

					sh.getRow(asset).getCell(2).setCellValue(txtBorrowedBy.getText());
					sh.getRow(asset).getCell(3).setCellValue(txtDate.getText());

					saveAsset();
					JOptionPane.showMessageDialog(null, "Device checked out successfully.", "Checked Out",
							JOptionPane.OK_OPTION);
					resetFields();
				} else {
					JOptionPane.showMessageDialog(null, "You need to enter the borrower's name.", "Checked Out",
							JOptionPane.OK_OPTION);
					txtBorrowedBy.requestFocus();
				}
			}
		});
		btnSubmit.setVisible(false);
		checkPanel.add(btnSubmit);
		checkPanel.add(btnCheckOut);

		btnCheckIn = new JButton("Check In");
		btnCheckIn.addActionListener(new ActionListener() {

			public void actionPerformed(ActionEvent e) {

				// Confirms that user is checking in device, if yes -- write fields, if no --
				// sets elements back to default
				int confirm = JOptionPane.showConfirmDialog(null, "Check in this device?", "Please confirm",
						JOptionPane.YES_NO_OPTION);

				if (confirm == 0) {

					int asset = getAsset(txtID.getText());

					sh.getRow(asset).getCell(2).setCellValue("Available");
					sh.getRow(asset).getCell(3).setCellValue("-");

					saveAsset();
					JOptionPane.showMessageDialog(null, "Device checked in successfully.", "Checked In",
							JOptionPane.OK_OPTION);

				} else {
					JOptionPane.showMessageDialog(null, "Check in cancelled.", "Checked In", JOptionPane.OK_OPTION);
				}

				resetFields();

			}
		});
		btnCheckIn.setVisible(false);
		checkPanel.add(btnCheckIn);

		btnCancel = new JButton("Cancel");
		btnCancel.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				// Sets elements to default
				resetFields();
			}
		});
		btnCancel.setVisible(false);
		checkPanel.add(btnCancel);

		// Added this just for the visual, thought it looked nice this way
		JTabbedPane tabbedPane = new JTabbedPane(JTabbedPane.TOP);
		contentPane.add(tabbedPane, BorderLayout.WEST);
	}

	// Method to get index of asset, returns 0 if not found
	public int getAsset(String id) {
		for (int i = 1; i <= sh.getLastRowNum(); i++) {
			int temp = (int) sh.getRow(i).getCell(0).getNumericCellValue();

			// System.out.println(temp);

			if (String.valueOf(temp).equals(id))
				return i;
		}

		return 0;
	}

	// Method to reset elements to default
	public void resetFields() {
		btnCheckIn.setVisible(false);
		btnCheckOut.setVisible(false);
		btnCancel.setVisible(false);
		btnSubmit.setVisible(false);
		txtBarcode.setEnabled(true);
		btnFind.setEnabled(true);
		txtBarcode.requestFocus();
		txtBorrowedBy.setText("");
		txtBorrowedBy.setEnabled(false);
		txtDate.setText("");
		txtID.setText("");
		txtType.setText("");
	}

	// Method to write to file
	public void saveAsset() {
		try {
			fos = new FileOutputStream("./inventory.xlsx");

			wb.write(fos);
			fos.flush();
			fos.close();

			resetFields();

		} catch (IOException e1) {
			// TODO Auto-generated catch block
			e1.printStackTrace();
		}
	}

}
