package com.gw.service;

import java.awt.Desktop;
import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import java.util.Map;

import javax.swing.JOptionPane;

import com.gw.ui.TestObjReportUI;
import com.gw.util.Constants;
import com.gw.util.MKSCommand;
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
			
			TestObjReportUI.logger.info("GET MKS connection completed!"); 
			TestObjReportUI.logger.info("Check the document ID completed!"); 
			ExcelUtil util = new ExcelUtil();
			TestObjReportUI.logger.info("start to export Test Suite report!");
			
			String exportType = TestObjReportUI.class.newInstance().exportType;
			String dept       = TestObjReportUI.class.newInstance().dept; 
			String filePath       = TestObjReportUI.class.newInstance().filePath; 
			TestObjReportUI.logger.info("current select type: "+ exportType);
			if (exportType==""){
				TestObjReportUI.logger.info("The export type is selected to be NULL! End of the program !!!");
				return;
			}
		    util.exportReport(tsIds, cmd, filePath,exportType);
//			根据sessionid查询测试结果 导出到 Excel模板 lxg
//		    util.getSessionIdByresultExportReport(tsIds, cmd, filePath,exportType);
//		    util.exportReport(tsIds, cmd, filePath,exportType);

		} catch (Exception e) {
			JOptionPane.showMessageDialog(TestObjReportUI.contentPane, e.getMessage(), "Error",
					JOptionPane.ERROR_MESSAGE);
			TestObjReportUI.logger.info("Error: " + e.getMessage());
			e.printStackTrace();
			success = false;
		} finally {
			if(success) {
				JOptionPane.showMessageDialog(TestObjReportUI.contentPane, "Success!");
			}
			try {
				cmd.release();
				TestObjReportUI.logger.info("cmd release!");
			} catch (IOException e) {
				TestObjReportUI.logger.info("Error: " + e.getMessage());
			} finally {
				System.exit(0);
			}
		}
	}
}
