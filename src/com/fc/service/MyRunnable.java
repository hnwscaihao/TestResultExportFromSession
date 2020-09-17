package com.fc.service;

import java.awt.Desktop;
import java.io.File;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.*;

import javax.swing.JOptionPane;

import com.alibaba.fastjson.JSONArray;
import com.alibaba.fastjson.JSONObject;
import com.fc.ui.ExportApplicationUI;
import com.fc.util.AnalysisXML;
import com.fc.util.Constants;
import com.fc.util.GenerateXmlUtil;
import com.fc.util.MKSCommand;
import com.mks.api.response.APIException;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;

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
			JSONArray ExData = new JSONArray(); //导出数据总json数组
			AnalysisXML analysisXML = new AnalysisXML();

			ExcelUtil util = new ExcelUtil();
			ExportApplicationUI.logger.info("选中的名字："+documentName);
			List<String> casefile = analysisXML.resultXmlType(documentName,"type","Test Case","field");
			ExportApplicationUI.logger.info("获取到的case字段"+casefile);
			List<Map<String,String>> lm = cmd.queryIssueByQuery(casefile,"((field[Document ID]="+tsIds.get(0)+"))");
			ExportApplicationUI.logger.info("获取到"+lm.size()+"条 Test Case");
			ExportApplicationUI.logger.info("查询的数据:" + lm);
			Map<String,String> allName = analysisXML.resultXmlPP(documentName);

			List<List<String>> listHeaders =  new ArrayList<>();//标题
			List<String> needMoreWidthField = new ArrayList<>();
			String name = "text";
			List< CellRangeAddress > merges = new ArrayList<>(); //跨行跨列
			getHead(listHeaders,merges);
			ExportApplicationUI.logger.info("获取的标题 :" + listHeaders);
			List<List<Object>> datas = new ArrayList<>();//导出数据集合

			for(Map<String,String> map : lm){
				List<Object> list= new ArrayList<>();
				for(int i=0;i<listHeaders.get(listHeaders.size()-1).size();i++){
					String almfield = allName.get(listHeaders.get(listHeaders.size()-1).get(i).toString());
					Object o = map.get(almfield);

					if(o == null){
						o = "";
					}
					if("Validates".equals(almfield) && !"".equals(o)){
						setR(o,map);
					}
					if("Blocked By".equals(almfield) && !"".equals(o)){
						setD(o,map);
					}
//					ExportApplicationUI.logger.info("每行数据:" + o);
					list.add(o);
				}
				datas.add(list);
			}

			Workbook wookbook  = new GenerateXmlUtil().exportComplexExcel(listHeaders,  datas, needMoreWidthField, name,merges , "");
			SimpleDateFormat format = new SimpleDateFormat("YYYYMMddHHmmss");
			ExcelUtil.outputFromwrok(filePath+"\\Test Suite " +format.format(new Date())+".xls", wookbook);

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

	/**
	 * 如果有需求id则查询需求赋值
	 */
    public void setR(Object id,Map<String,String> map){
		//获取需求内容
		List<String> casefile = new AnalysisXML().resultXmlType(documentName,"type","Requirements","field");
		try {
			String [] ids = id.toString().split(",");
			List<String> lids = new ArrayList<>();
			for(int i= 0;i<ids.length;i++){
				lids.add(ids[i]);
			}
			List<Map<String,String>> ml = cmd.getItemByIds(lids,casefile);
			Map<String,String> mr = new HashMap<>();

			for(String s:casefile){
				map.put(s,getvalues(ml,s));
			}
		} catch (APIException e) {
			e.printStackTrace();
		}
	}
	public String getvalues(List<Map<String,String>> ml,String key){
    	String str = "";
		for(Map<String,String> m:ml){
			for(String s:m.keySet()){
				if(s.equals(key)){
					ExportApplicationUI.logger.info("获取的标题 :" + m.get(s));
					str += m.get(s)+",";
				}
			}
		}
		if(str.length()>0){
			str = str.substring(0,str.length()-1);
		}
		return str;
	}
	/**
	 * 如果有缺陷id则查询缺陷赋值
	 */
	public void setD(Object id,Map<String,String> map){
		List<String> casefile = new AnalysisXML().resultXmlType(documentName,"type","Defect","field");
		try {
			String [] ids = id.toString().split(",");
			List<String> lids = new ArrayList<>();
			for(int i= 0;i<ids.length;i++){
				lids.add(ids[i]);
			}
			List<Map<String,String>> ml = cmd.getItemByIds(lids,casefile);
			for(String s:casefile){
				map.put(s,getvalues(ml,s));
			}
		} catch (APIException e) {
			e.printStackTrace();
		}
	}

	/**
	 * 标题和跨行处理
	 * @param listHeaders
	 * @param merges
	 */
	public void getHead(List<List<String>> listHeaders ,List< CellRangeAddress > merges){

		AnalysisXML analysisXML = new AnalysisXML();
		//文档标题
		List<String> DocHead = new ArrayList<>();
		//最上面夸行标题
		List<String> head1 = new ArrayList<>();
		List<String> head2 = new ArrayList<>();
		List<String> head3 = new ArrayList<>();
		List<String> head4 = new ArrayList<>();
		List<String> head5 = new ArrayList<>();
		List<String> head6 = new ArrayList<>();
		List<String> head7 = new ArrayList<>();
		List<String> head8 = new ArrayList<>();
		List<String> head9 = new ArrayList<>();
		List<String> head10 = new ArrayList<>();
		List<String> head11 = new ArrayList<>();
		List<String> head12 = new ArrayList<>();
		List<String> head13 = new ArrayList<>();
		List<String> head14 = new ArrayList<>();
		List<String> head15 = new ArrayList<>();
		List<String> head16 = new ArrayList<>();
		List<String> head17 = new ArrayList<>();
		List<String> head18 = new ArrayList<>();
		List<String>  h1s = analysisXML.resultXmlType(documentName,"title","h1","name");
		List<String>  h2s = analysisXML.resultXmlType(documentName,"title","h2","name");
		Map<String, String> h1data = new HashMap<>();
		try {
			h1data = cmd.getItemByIds(tsIds, Arrays.asList("ID","Text", "Document Short Title","Assigned User","Tests For")).get(0);
		} catch (APIException e) {
			e.printStackTrace();
		}
		for(int i=0;i<h1s.size();i++){
			DocHead.add(h1s.get(i));
			if(i==0){
				head1.add("Test Suite ID 测试用例组ID : " + h1data.get("ID"));
				head2.add("Test Suite Description 测试用例组描述 : " + h1data.get("Text"));
				head3.add("Test Suite Type 测试用例组类型 : " + h1data.get("Document Short Title"));
				head4.add("Test Suite Change Date 测试用例组变更时间 : " );
				head5.add("Test Suite Change Log 测试用例组变更描述记录 : ");
				head6.add("Test Suite Change by 测试用例组变更执人 : " + h1data.get("Assigned User"));
				head7.add("");
				head8.add("");
				head9.add("");
				head10.add("");
				head11.add("");
				head12.add("");
				head13.add("");
				head14.add("");
				head15.add("");
				head16.add("");
				head17.add("");
				head18.add("");
			}else {
				head1.add("");
				head2.add("");
				head3.add("");
				head4.add("");
				head5.add("");
				head6.add("");
				head7.add("");
				head8.add("");
				head9.add("");
				head10.add("");
				head11.add("");
				head12.add("");
				head13.add("");
				head14.add("");
				head15.add("");
				head16.add("");
				head17.add("");
				head18.add("");
			}
		}
		merges.add(new CellRangeAddress(0,0,0,h1s.size()-1));
		merges.add(new CellRangeAddress(1,1,0,h1s.size()-1));
		merges.add(new CellRangeAddress(2,2,0,h1s.size()-1));
		merges.add(new CellRangeAddress(3,3,0,h1s.size()-1));
		merges.add(new CellRangeAddress(4,4,0,h1s.size()-1));
		merges.add(new CellRangeAddress(5,5,0,h1s.size()-1));
		merges.add(new CellRangeAddress(6,6,0,h1s.size()-1));
		merges.add(new CellRangeAddress(7,7,0,h1s.size()-1));
		merges.add(new CellRangeAddress(8,8,0,h1s.size()-1));
		merges.add(new CellRangeAddress(9,9,0,h1s.size()-1));
		merges.add(new CellRangeAddress(10,10,0,h1s.size()-1));
		merges.add(new CellRangeAddress(11,11,0,h1s.size()-1));
		merges.add(new CellRangeAddress(12,12,0,h1s.size()-1));
		merges.add(new CellRangeAddress(13,13,0,h1s.size()-1));
		merges.add(new CellRangeAddress(14,14,0,h1s.size()-1));
		merges.add(new CellRangeAddress(15,15,0,h1s.size()-1));
		merges.add(new CellRangeAddress(16,16,0,h1s.size()-1));
		merges.add(new CellRangeAddress(17,17,0,h1s.size()-1));
//		merges.add(new CellRangeAddress(18,18,0,h1s.size()-1));

		String[] sessins = h1data.get("Tests For").split(",");
		for(int s =0 ;s<sessins.length;s++){
			Map<String,String> sMap= null;
			try {
				sMap = cmd.getItemByIds(Arrays.asList(sessins[s]),
						Arrays.asList("ID","Summary","Type","Actual Start Date","Actual End Date","Assigned User",
								"Planned Count","Pass Percentage","Pass Count","Fail Count","Spawned By")).get(0);
			} catch (APIException e) {
				e.printStackTrace();
			}
			for(int i=0;i<h2s.size();i++){
				if("Test Session".equals(sMap.get("Type"))){
					DocHead.add(h2s.get(i));
					if(i==0){
						head1.add("Test Session ID 测试任务ID No."+(i+1) +": " +sMap.get("ID"));
						head2.add("Test Session Abstract  测试任务描述: "+sMap.get("Summary"));
						head3.add("Test Session Start Date/End Date 测试任务开始/结束时间 : "+sMap.get("Actual Start Date") +"/"+sMap.get("Actual End Date"));
						head4.add("Test Session Result 测试任务结果");
						head5.add("Test Engineer 测试人员,测试执行人: "+sMap.get("Assigned User"));
						head6.add("Test Effort 测试任务耗费工时");
						head7.add("Test Environment 测试环境");
						head8.add("Test Type 测试类型");
						head9.add("Test Equipment 测试设备");
						head10.add("Test Software 测试软件");
						head11.add("Hardware Version 测试对象硬件版本号");
						head12.add("Software Version 测试对象软件版本号");
						head13.add("Calibration File Version 测试对象标定文件版本号");
						head14.add("Test Case NO 执行的测试用例数量： "+sMap.get("Planned Count"));
						head15.add("Test Case Passed NO 测试用例通过数： "+sMap.get("Pass Count"));
						head16.add("Test Case Failed NO 测试用例不通过数： "+sMap.get("Fail Count"));
						head17.add("Test Passing Rate 测试用例通过率： "+sMap.get("Pass Percentage"));
						int qxs = 0;
						if(sMap.get("Spawned By") != null && "".equals(sMap.get("Spawned By"))){
							qxs = sMap.get("Spawned By").split(",").length;
						}
						head18.add("Defects NO 测试发现缺陷数量:"+qxs);
					}else {
						head1.add("");
						head2.add("");
						head3.add("");
						head4.add("");
						head5.add("");
						head6.add("");
						head7.add("");
						head8.add("");
						head9.add("");
						head10.add("");
						head11.add("");
						head12.add("");
						head13.add("");
						head14.add("");
						head15.add("");
						head16.add("");
						head17.add("");
						head18.add("");
					}
				}
			}
			int qsh = 0;
			if(s==0){
				qsh = h1s.size();
			}else {
				qsh = h1s.size()+(h2s.size()*s);
			}
			int jsh = qsh +h2s.size()-1;
			merges.add(new CellRangeAddress(0,0,qsh,jsh));
			merges.add(new CellRangeAddress(1,1,qsh,jsh));
			merges.add(new CellRangeAddress(2,2,qsh,jsh));
			merges.add(new CellRangeAddress(3,3,qsh,jsh));
			merges.add(new CellRangeAddress(4,4,qsh,jsh));
			merges.add(new CellRangeAddress(5,5,qsh,jsh));
			merges.add(new CellRangeAddress(6,6,qsh,jsh));
			merges.add(new CellRangeAddress(7,7,qsh,jsh));
			merges.add(new CellRangeAddress(8,8,qsh,jsh));
			merges.add(new CellRangeAddress(9,9,qsh,jsh));
			merges.add(new CellRangeAddress(10,10,qsh,jsh));
			merges.add(new CellRangeAddress(11,11,qsh,jsh));
			merges.add(new CellRangeAddress(12,12,qsh,jsh));
			merges.add(new CellRangeAddress(13,13,qsh,jsh));
			merges.add(new CellRangeAddress(14,14,qsh,jsh));
			merges.add(new CellRangeAddress(15,15,qsh,jsh));
			merges.add(new CellRangeAddress(16,16,qsh,jsh));
			merges.add(new CellRangeAddress(17,17,qsh,jsh));
		}



		listHeaders.add(head1);
		listHeaders.add(head2);
		listHeaders.add(head3);
		listHeaders.add(head4);
		listHeaders.add(head5);
		listHeaders.add(head6);
		listHeaders.add(head7);
		listHeaders.add(head8);
		listHeaders.add(head9);
		listHeaders.add(head10);
		listHeaders.add(head11);
		listHeaders.add(head12);
		listHeaders.add(head13);
		listHeaders.add(head14);
		listHeaders.add(head15);
		listHeaders.add(head16);
		listHeaders.add(head17);
		listHeaders.add(head18);
		listHeaders.add(DocHead);
	}

	public static void main(String[] s){
		List<List<String>> listHeaders = new ArrayList<>();
		List<String> oneHeader = new ArrayList<>();
		List<String> TwoHeader = new ArrayList<>();
		oneHeader.add("1.1");
		oneHeader.add("1.2");
		oneHeader.add("1.3");

		TwoHeader.add("2.1");
		TwoHeader.add("");
		TwoHeader.add("");
		listHeaders.add(oneHeader);
		listHeaders.add(TwoHeader);


		List<List<Object>> datas = new ArrayList<>();
		List<String> needMoreWidthField = new ArrayList<>();
		String name = "text";
		List< CellRangeAddress > merges = new ArrayList<>();
		merges.add(new CellRangeAddress(1,1,0,2));
		Workbook wookbook  = new GenerateXmlUtil().exportComplexExcel(listHeaders,  datas, needMoreWidthField, name,merges , "");
		SimpleDateFormat format = new SimpleDateFormat("YYYYMMddHHmmss");
		ExcelUtil.outputFromwrok("E:\\ " +format.format(new Date())+".xls", wookbook);


	}
}
