package com.main.excel;

import java.awt.*;
import java.awt.event.*;

import javax.swing.*;

import com.main.excel.gui.CreateRunChart;

public class JavaWaitCursorExample implements ActionListener {
	private JFrame frame;
	private String defaultButtonText = "Show Wait Cursor";
	private JButton button = new JButton(defaultButtonText);
	private Cursor waitCursor = new Cursor(Cursor.WAIT_CURSOR);
	private Cursor defaultCursor = new Cursor(Cursor.DEFAULT_CURSOR);
	private boolean waitCursorIsShowing;

	public static void main(String[] args) {
		new JavaWaitCursorExample();
	}

	public JavaWaitCursorExample() {
		button.addActionListener(this);
		frame = new JFrame("Java Wait Cursor Example");
		frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
		frame.getContentPane().add(button);
		frame.setPreferredSize(new Dimension(400, 300));
		frame.pack();
		frame.setLocationRelativeTo(null);
		frame.setVisible(true);
	}

	public void actionPerformed(ActionEvent evt) {
		CreateRunChart createRunChart = new CreateRunChart();
		waitCursorIsShowing=createRunChart.getExcel("1430312362558","R");
		if (waitCursorIsShowing) {
			waitCursorIsShowing = false;
			button.setText(defaultButtonText);
			frame.setCursor(defaultCursor);
		} else {
			waitCursorIsShowing = true;
			button.setText("Back to Default Cursor");
			frame.setCursor(waitCursor);
		}
	}
}