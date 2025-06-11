package com.main.excel.gui;

import java.awt.BorderLayout;
import java.awt.Frame;
import java.awt.GridBagConstraints;
import java.awt.GridBagLayout;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;

import javax.swing.JButton;
import javax.swing.JDialog;
import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JOptionPane;
import javax.swing.JPanel;
import javax.swing.JTextField;
 
public class RunDialog extends JDialog {
	
	private static final long serialVersionUID = 1L;
	final static JFrame frame = new JFrame("JDialog Demo");
	private JTextField runID;
    private JTextField type;
    private JLabel runIDLabel;
    private JLabel typeLabel;
    private JButton btnRun;
    private JButton btnCancel;
    private boolean succeeded;
    
    public static void main(String[] args) {	
    	
        RunDialog loginDlg = new RunDialog(frame);
        loginDlg.setVisible(true);
    }
    
    public RunDialog(Frame parent) {
        super(parent, "Run Details", true);
        JPanel panel = new JPanel(new GridBagLayout());
        GridBagConstraints cs = new GridBagConstraints();
        cs.fill = GridBagConstraints.HORIZONTAL;
 
        runIDLabel = new JLabel("Run/Batch ID: ");
        cs.gridx = 0;
        cs.gridy = 0;
        cs.gridwidth = 1;
        panel.add(runIDLabel, cs);
 
        runID = new JTextField(20);
        cs.gridx = 1;
        cs.gridy = 0;
        cs.gridwidth = 2;
        panel.add(runID, cs);
        
        typeLabel = new JLabel("Type Batch(B)/Run(R): ");
        cs.gridx = 0;
        cs.gridy = 1;
        cs.gridwidth = 1;
        panel.add(typeLabel, cs);
 
        type = new JTextField(20);
        cs.gridx = 1;
        cs.gridy = 1;
        cs.gridwidth = 2;
        panel.add(type, cs);
        
        btnRun = new JButton("Run");
        btnRun.addActionListener(new ActionListener() {
            public void actionPerformed(ActionEvent e) {
                if (validateText()) {            	
                	CreateRunChart createRunChart = new CreateRunChart();
                	succeeded=createRunChart.getExcel(getRunID(), getBatchType());
                	if(succeeded) {
                	JOptionPane.showMessageDialog(RunDialog.this,
                            "You have successfully Extracted the Run Details.",
                            "Run",
                            JOptionPane.INFORMATION_MESSAGE);
                	}
                    dispose();
                } else {
                    JOptionPane.showMessageDialog(RunDialog.this,
                            "Invalid Entry",
                            "Run",
                            JOptionPane.ERROR_MESSAGE);
                    succeeded = false;
 
                }
            }
        });
        btnCancel = new JButton("Cancel");
        btnCancel.addActionListener(new ActionListener() {
            public void actionPerformed(ActionEvent e) {
                dispose();
            }
        });
        JPanel bp = new JPanel();
        bp.add(btnRun);
        bp.add(btnCancel);
 
        getContentPane().add(panel, BorderLayout.CENTER);
        getContentPane().add(bp, BorderLayout.PAGE_END);
 
        pack();
        setResizable(false);
        setLocationRelativeTo(parent);
    }
 
    public String getRunID() {
        return runID.getText().trim();
    }
   
 
    public boolean isSucceeded() {
        return succeeded;
    }
    
    public String getBatchType() {
        return type.getText().trim().toUpperCase();
    }
    
    protected boolean validateText() {
    	if(getRunID()==null || getRunID().trim().equals("")) {
    		return false;
    	}
    	else if(getBatchType()==null || getBatchType().trim().equals(""))
    		return false;
    	else
    		return true;
    }
}