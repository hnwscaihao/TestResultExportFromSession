package com.fc.util;


import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Properties;

public class ConfigFiled {
    /**
     * @Description
     * @Author  liuxiaoguang
     * @Date   2020/8/25 15:12
     * @Param  [path, fileds]
     * @Return      java.lang.String
     * @Exception   获取配置文件字段 v1 配置文件名字 v2 配置文件中字段
     */
    public static Map<String,String> getPath1(String path, List<String> fileds) {

        Map<String,String> map = new HashMap<>();
        try {
            InputStream in = ConfigFiled.class.getClassLoader().getResourceAsStream(path);
            Properties prop = new Properties();
            prop.load(in);

            for(int i=0;i<fileds.size();i++){
                String s = fileds.get(i);
                map.put(s,prop.getProperty(s));
            }
        } catch (FileNotFoundException e) {
            System.out.println("properties文件路径书写有误，请检查！");
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }

        return map;
    }

    public static void main(String[] str){
//        Map<String,String>  map = getPath1("ConfigFiled.properties",Arrays.asList("loginName","passWord"));
//       new  ExcelUtil().getListByExcel();
    }
}
