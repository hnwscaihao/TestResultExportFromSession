package com.fc.service;

import java.awt.Desktop;
import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import java.util.Map;

import javax.swing.JOptionPane;

import com.fc.ui.ExportApplicationUI;
import com.fc.util.Constants;
import com.fc.util.MKSCommand;
@SuppressWarnings("all")
public class MyRunnable implements Runnable {

	public MKSCommand cmd;
	public List<String> tsIds = new ArrayList<>();//存放获取到的ID集合
	public String filePath;
	public static String documentName;
	public MyRunnable() {
		super();
	}
	
	@Override
	public void run() {
		boolean success = true;
		try {
//			String userHome = TestObjReportUI.ENVIRONMENTVAR.get(Constants.USER_HOME); //获取用户目录
//			if (userHome == null || userHome.length() == 0) {
//				userHome = TestObjReportUI.ENVIRONMENTVAR.get("USERPROFILE");//如果没获取到 手动再次获取
//			}
//			if (userHome == null || userHome.length() == 0) {
//				userHome = "C:\\Users\\" + TestObjReportUI.ENVIRONMENTVAR.get(Constants.USERNAME);//再次获取用户目录
//			}
//			filePath = input.toString();
//			if (!input.exists()) { //如果不存在就去创建
//				input.mkdirs();
//			}
			
			ExportApplicationUI.logger.info("GET MKS connection completed!"); 
			ExportApplicationUI.logger.info("Check the document ID completed!"); 
			ExcelUtil util = new ExcelUtil();
			ExportApplicationUI.logger.info("start to export Test Suite report!");
			
			String exportType = ExportApplicationUI.class.newInstance().exportType;
			String dept       = ExportApplicationUI.class.newInstance().dept; 
			String filePath       = ExportApplicationUI.class.newInstance().filePath; 
			ExportApplicationUI.logger.info("current select type: "+ exportType);
			if (exportType==""){
				ExportApplicationUI.logger.info("The export type is selected to be NULL! End of the program !!!");
				return;
			}
		    util.exportReport(tsIds, cmd, filePath,exportType);

		} catch (Exception e) {
			JOptionPane.showMessageDialog(ExportApplicationUI.contentPane, e.getMessage(), "Error",
					JOptionPane.ERROR_MESSAGE);
			ExportApplicationUI.logger.info("Error: " + e.getMessage());
			e.printStackTrace();
			success = false;
		} finally {
			if(success) {
				JOptionPane.showMessageDialog(ExportApplicationUI.contentPane, "Success!");
			}
			try {
				cmd.release();
				ExportApplicationUI.logger.info("cmd release!");
			} catch (IOException e) {
				ExportApplicationUI.logger.info("Error: " + e.getMessage());
			} finally {
				System.exit(0);
			}
		}
	}
}
