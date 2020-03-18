package com.fc.ui;

import java.awt.EventQueue;
import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Date;
import java.util.List;
import java.util.Map;
import java.util.Set;
import java.util.Timer;
import java.util.TimerTask;
import java.util.concurrent.Executors;
import java.util.concurrent.ScheduledExecutorService;
import java.util.concurrent.TimeUnit;

import javax.swing.JFrame;
import javax.swing.JOptionPane;
import javax.swing.JPanel;
import javax.swing.JProgressBar;
import javax.swing.UIManager;
import javax.swing.border.EmptyBorder;
import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.parsers.ParserConfigurationException;

import org.apache.log4j.Logger;
import org.w3c.dom.Document;
import org.w3c.dom.Element;
import org.w3c.dom.NodeList;
import org.xml.sax.SAXException;

import com.fc.service.ExcelUtil;
import com.fc.service.MyRunnable;
import com.fc.util.Constants;
import com.fc.util.MKSCommand;

import javax.swing.DefaultComboBoxModel;
import javax.swing.JComboBox;
import javax.swing.JFileChooser;

import java.awt.Label;
import java.awt.Font;
import java.awt.Color;
import java.awt.Button;
import java.awt.event.ActionListener;
import java.awt.event.ItemEvent;
import java.awt.event.ItemListener;
import java.awt.event.ActionEvent;
import javax.swing.JLabel;
import javax.swing.border.EtchedBorder;
import javax.swing.event.ChangeEvent;
import javax.swing.event.ChangeListener;
import javax.swing.JButton;
import javax.swing.SwingConstants;
@SuppressWarnings("all")
public class ExportApplicationUI extends JFrame {
	private static final long serialVersionUID = 1L;
	public static JPanel contentPane;
	public static final Map<String, String> ENVIRONMENTVAR = System.getenv();
	public static MKSCommand cmd;
	public static List<String> tsIds = new ArrayList<>();
	public static List<String> caseIDList = new ArrayList<>();
	public static final Logger logger = Logger.getLogger(ExportApplicationUI.class);
	public static Map<String,String> sessionMap ; // 直接查询session信息，保存起来，防止多次查询
	JProgressBar progressBar;
	private static final String EMPTY_IMPORT_TYPE = "Please Select a Type";
	public static String exportType;
	private static String DOCUMENT_TYPE ; 
	private boolean start;
	public static String dept;
	private MyRunnable run = new MyRunnable();
	private Thread thread = new Thread();// 查询线程
	private static JComboBox comboBox;
	private JButton button;
	private JLabel label_1;
	private JLabel lblNewLabel;
	private JComboBox comboBox_1;
	public  static int MIN_PROGRESS = 0;
	public  static int MAX_PROGRESS = 8;
	public static int currentProgress = MIN_PROGRESS;
	public static ScheduledExecutorService pool;
	public static boolean isParseSuccess = false;//判断导出是否结束
	public static String trueType;
	public static  String filePath;
	public static boolean ispause=false;
	public static boolean isStart=false;
	public static boolean chooseRight = false;
	private ExcelUtil excelUtil = new ExcelUtil();
	private static List<String> typeList = null;
	
	private static String documentName ;
	/**
	 * Launch the application.
	 */
	public static void main(String[] args) {
		EventQueue.invokeLater(new Runnable() {
			public void run() {
				try {
					MKSCommand m = new MKSCommand();
					initMksCommand();//初始化MKSCommand中的参数，并获得连接
					UIManager.setLookAndFeel(UIManager.getSystemLookAndFeelClassName());//设置与本机适配的swing样式
					ExportApplicationUI frame = new ExportApplicationUI();
					frame.setVisible(true);
					frame.setLocationRelativeTo(null);
					frame.setValueForPick();//选择部门设置数据
					getSelectedIdList();//获取到当前选中的id添加进集合Ids集合
				} catch (Exception e) {
					JOptionPane.showMessageDialog(contentPane, e.getMessage());
					System.exit(0);
				}
			}
		});
	}

	/**
	 * 开始导出
	 * @throws Exception
	 */
	public void startExport() throws Exception {
		run.documentName = documentName;
		run.tsIds = caseIDList;
		run.cmd = cmd;
		thread = new Thread(run);
		thread.start();//
	}
	
	public void  updateProgress (){
	    pool = Executors.newScheduledThreadPool(1);
	    pool.scheduleAtFixedRate(new Runnable() {
			@Override
			public void run() {
				currentProgress++;
			        if (currentProgress > MAX_PROGRESS && !isParseSuccess) {
			            currentProgress = MIN_PROGRESS;
			        }
			        if(isParseSuccess) {
			        	getProgressBar().setValue(MAX_PROGRESS);
			        	pool.shutdown();
			        }
			        getProgressBar().setValue(currentProgress);
			}
			}, 
			1, 
			1,
			TimeUnit.SECONDS);
	}
	
		
	/**
	 * 解析配置文件为Pick设置数据
	 */
	public void setValueForPick() {
		try {
			typeList = excelUtil.parseXML(null);//解析xml文件
			typeList.add(0, "Please select a type");
		} catch (Exception e) {
			JOptionPane.showMessageDialog(this, "Could Not Parse XML Config, Please Contant Adminstrator!");
		}
		comboBox.setModel(new DefaultComboBoxModel<String>(typeList.toArray(new String[typeList.size()])));
		if(DOCUMENT_TYPE != null && !"".equals(DOCUMENT_TYPE) && typeList.contains(DOCUMENT_TYPE)){
			comboBox.setSelectedItem(DOCUMENT_TYPE);
		}
	}
	
	
	
	/**
	 * Create the frame.
	 */
	public ExportApplicationUI() {
		setAutoRequestFocus(false);
		setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
		setTitle("Export Excel Report");
		setResizable(false);
		setBounds(100, 100, 820, 318);
		contentPane = new JPanel();
		contentPane.setForeground(Color.WHITE);
		contentPane.setBorder(new EmptyBorder(5, 5, 5, 5));
		setContentPane(contentPane);
		contentPane.setLayout(null);
		progressBar = new JProgressBar();
		progressBar.setString("Exporting, please wait!");
		progressBar.setBounds(82, 213, 539, 24);
        // 设置当前进度值
        getProgressBar().setValue(0);
        getProgressBar().setMinimum(MIN_PROGRESS);
        getProgressBar().setMaximum(MAX_PROGRESS);
		contentPane.add(progressBar);
       //选择excel模板
		comboBox = new JComboBox();
		comboBox.setBounds(82, 51, 540, 36);
		contentPane.add(comboBox);
		
		button = new JButton("Export");
		
		button.addActionListener(new ActionListener() {
			
			public void actionPerformed(ActionEvent e) {
				exportType = comboBox.getSelectedItem().toString();
				trueType =exportType;
				String selectedIndex = comboBox.getSelectedItem().toString();
				if(selectedIndex.equals("Please select a type")||selectedIndex=="selectedIndex"){
					JOptionPane.showConfirmDialog(contentPane, "Select the correct type for export!!");//导出类型
					return;
				}
				if (label_1.getText().equals("< Please select an export path />")) {
					
					JOptionPane.showConfirmDialog(contentPane, "Please select a folder as the path to export!!");//导出路径
					return;
				}else{
					updateProgress();
				}
				//开始导出！
				try {
					startExport();
					button.setEnabled(false); 
				} catch (Exception e1) {
					JOptionPane.showMessageDialog(contentPane, e1.getMessage());
					System.exit(0);
				}
			}
		});
		button.setForeground(Color.BLACK);
		button.setBounds(662, 207, 137, 30);
		contentPane.add(button);
		label_1 = new JLabel("< Please select an export path />");
		label_1.setForeground(Color.BLACK);
		label_1.setBorder(new EtchedBorder());
		label_1.setBounds(82, 132, 539, 37);
		contentPane.add(label_1);
		
		JButton button_1 = new JButton("Browse");
		button_1.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				ExportApplicationUI.this.browseDocAction();//
			}
		});
		button_1.setBounds(662, 136, 137, 29);
		contentPane.add(button_1);
		
		lblNewLabel = new JLabel("[First of all, please choose your own department.]");
		lblNewLabel.setForeground(Color.RED);
		lblNewLabel.setBounds(82, 15, 480, 21);
		
		comboBox_1 = new JComboBox();
		comboBox_1.setBounds(82, 51, 100, 37);

		JLabel label_2 = new JLabel("*");
		label_2.setVerticalAlignment(SwingConstants.TOP);
		label_2.setForeground(Color.RED);
		label_2.setBounds(636, 59, 9, 21);
		contentPane.add(label_2);

		JLabel label_3 = new JLabel("*");
		label_3.setVerticalAlignment(SwingConstants.TOP);
		label_3.setForeground(Color.RED);
		label_3.setBounds(183, 59, 9, 21);

	}
	/**
	 * 浏览文件操作
	 */
	protected void browseDocAction() {
		logger.info("Start to load file of import"); 
		JFileChooser fc = new JFileChooser();
		fc.setDialogTitle("选择一个路径");
		fc.setAcceptAllFileFilterUsed(true);
		fc.setMultiSelectionEnabled(false);
		fc.setFileSelectionMode(JFileChooser.FILES_AND_DIRECTORIES );
		int returnVal = fc.showOpenDialog(this);
		if (returnVal == 0) {
			File input = fc.getSelectedFile();
			filePath = input.toString();
			logger.info("selected file:" + input);
			if(!input.isDirectory() && !input.toString().endsWith("xls") ){//生成03版本，所以是xls后缀
				JOptionPane.showMessageDialog(contentPane, "Please Choose a folder or a XLS file");
				chooseRight = false;
			}else{
				label_1.setText("path:"+input.toString());
				chooseRight = true;
			}
		}
	}
	
	public JProgressBar getProgressBar() {
		return progressBar;
	}


	public void setProgressBar(JProgressBar progressBar) {
		this.progressBar = progressBar;
	}

	/**
	 * 初始化MKSCommand中的参数，并获得连接
	 */
	public static void initMksCommand() {
		try {
			String host = ExportApplicationUI.ENVIRONMENTVAR.get(Constants.MKSSI_HOST);
			if(host==null || host.length()==0) {
//				host = "192.168.6.130";
				host = "192.168.6.130";
//				host = "192.168.229.133";
//				host = "10.255.33.189";
			}
			String portStr = ENVIRONMENTVAR.get(Constants.MKSSI_PORT);
			Integer port = portStr!=null && !"".equals(portStr)? Integer.valueOf(portStr) : 7001;
			String defaultUser = ENVIRONMENTVAR.get(Constants.MKSSI_USER);
			String pwd = "";
			if(defaultUser == null || "".equals(defaultUser) ){
//				defaultUser = "GW00165407";
//				pwd = "123369";
				defaultUser = "admin";
				pwd = "admin";
			}
			logger.info("host:" + host+"; defaultUser:"+defaultUser+"; pwd:"+pwd);
			cmd = new MKSCommand(host, 7001, defaultUser, pwd, 4, 16);
		} catch (Exception e) {
			JOptionPane.showMessageDialog(ExportApplicationUI.contentPane, "Can not get a connection!", "Message",
					JOptionPane.WARNING_MESSAGE);
			ExportApplicationUI.logger.info("Can not get a connection!");
			System.exit(0);
		}
	}

	/**
	 * 获取当前选中id的List集合
	 * @return
	 * @throws Exception
	 */
	public static List<String> getSelectedIdList() throws Exception {
		String issueCount = ExportApplicationUI.ENVIRONMENTVAR.get(Constants.MKSSI_NISSUE);
		ExportApplicationUI.logger.info("get issue count from environment:" + issueCount); 
		if (issueCount != null && issueCount.trim().length() > 0) { 
			for (int index = 0; index < Integer.parseInt(issueCount); index++) {
				String id = ExportApplicationUI.ENVIRONMENTVAR.get(String.format(Constants.MKSSI_ISSUE_X, index));
				ExportApplicationUI.logger.info("get the selection test Suite : " + id);
				tsIds.add(id);//获取到当前选中的id添加进集合Ids集合
			}
		} else {
			ExportApplicationUI.logger.info("No ID was obtained!!! :" + issueCount); 
		}
//		tsIds.add("54118");
//		tsIds.add("11206");
		if(tsIds.size()>1){
			JOptionPane.showMessageDialog(null, "暂时不支持多选！","错误",0);
			System.exit(0);
		}
		if (tsIds.size() > 0) {//如果选中的id集合不为空，通过id获取条目简要信息
			sessionMap = cmd.getItemByIds(tsIds, Arrays.asList("ID", "Type","Summary","tests")).get(0);
			List<String> notTSList = new ArrayList<>();
				DOCUMENT_TYPE = sessionMap.get("Type");
				String id = sessionMap.get("ID");
				documentName = sessionMap.get("Summary");
				String caseids = sessionMap.get("tests");
				if(typeList.contains(DOCUMENT_TYPE)){
					comboBox.setSelectedItem(DOCUMENT_TYPE);
				}
				if(!DOCUMENT_TYPE.equals("Test Session")){
					JOptionPane.showMessageDialog(null, "请选择正确的Test Session ID！","错误",0);
					System.exit(0);
				}
				if(caseids == null || caseids.length() == 0){
					JOptionPane.showMessageDialog(null, "当前Test Session无有效测试结果项！","错误",0);
					System.exit(0);
				}else{//将所有Test Case ID记录下来
					String[] caseIDStrs = caseids.split(","); 
					List<String> caseTempList = new ArrayList<>();
					for(String caseID : caseIDStrs){
						caseTempList.add(caseID);
					}
					List<Map<String,String>> itemByIds = cmd.getItemByIds(caseTempList, Arrays.asList("ID", "Type"));
					for(Map<String, String> caseMap : itemByIds){
						String type = caseMap.get("Type");
						String issueId = caseMap.get("ID");
						if("Test Suite".equals(type)){
							List<String> allContains = cmd.allContainID(issueId);
							if(allContains != null && !allContains.isEmpty()){
								for(String containId : allContains){
									caseIDList.add(containId);
								}
							}
						}else{
							caseIDList.add(issueId);
						}
					}
				}
				
			if (notTSList.size() > 0) {
//				throw new Exception("This item " + notTSList + " is not [ " + documentType + " ]! Please  select the right type!");
			} else {
				ExportApplicationUI.logger.info("get the selection Document : " + tsIds);
			}
		} else {
			throw new Exception("Please select the ID of a Document!");
		}
		return tsIds;
	}
}
