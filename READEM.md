# <center>Test Objective	Report导出</center>
#### 一、功能介绍
1. 导出报表为excel模板，模板配置参考FieldMapping.xml中的配置项说明;<br/>
2. 导出的数据涉及到Test Objective，Test Session，Test Suite，Test Case及其与Test Session执行测试所得到的测试结果，每种不同测试结果分别用不同颜色标注;<br/>
3. 用户可同时选择多个Test Objective进行导出，当选择数据不是Test Objective时，会有提示选择错误;<br/>
4. 导出报表完毕后，会自动打开所在文件夹，供用户快速找到导出excel文件.
5. logger日志文件见：${user.home}/ImpAndExpLogs/TestObjReport.log

#### 二、操作步骤
1. 在Integrity界面中选择一条或几条Test Objective<br/>
2. 在Custom中选择Test Objective/Session Report菜单进行导出<br/>

2020-5-10
