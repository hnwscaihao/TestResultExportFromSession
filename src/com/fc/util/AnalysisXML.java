package com.fc.util;

import com.alibaba.fastjson.JSON;
import com.alibaba.fastjson.JSONArray;
import com.alibaba.fastjson.JSONObject;
import org.apache.commons.logging.Log;
import org.apache.commons.logging.LogFactory;
import org.dom4j.Attribute;
import org.dom4j.Document;
import org.dom4j.Element;
import org.dom4j.io.SAXReader;

import java.io.InputStream;
import java.util.*;

/*
*Dom4j解析xml
 */
public class AnalysisXML {

    private static final Log log = LogFactory.getLog(AnalysisXML.class);

    public static Map<String,String> relationshipMap = new HashMap<String,String>();
    
    private final static String SPLIT_FLAG = "|q|q|";
    
    public String resultCategory(String type1,String type2){
        String s = "";
       //1.创建Reader对象
       SAXReader reader = new SAXReader();
       //2.加载xml
       Document document = null;
        SAXReader saxReader = new SAXReader();
       try {

           //返回读取指定资源的输入流
           InputStream in = AnalysisXML.class.getClassLoader().getSystemResourceAsStream("Category.xml");
//           writeToLocal(pathxml,in);
           document = reader.read(in);

           //3.获取根节点
           Element rootElement = document.getRootElement();
           Iterator iterator = rootElement.elementIterator();
           while (iterator.hasNext()){
               Element stu = (Element) iterator.next();
               List<Attribute> attributes = stu.attributes();
//               System.out.println("======获取属性值======");
               if(stu.attribute("swr").getValue().equals(type1)){
                   Iterator iterator1 = stu.elementIterator();

                   while (iterator1.hasNext()){
                       Element stuChild = (Element) iterator1.next();
                       if(stuChild.attribute("categorySWR").getValue().equals(type2)){
                           s = stuChild.attribute("categoryalm").getValue();
                       }
                   }
               }
           }
       } catch (Exception e) {
           e.printStackTrace();
       }
       return s;
   }

    /** 
     * 获取关系字段
     * @param currType
     * @param targetType
     * @return
     */
    public static String getRelationshipField(String currType, String targetType){
    	String relationshipField = relationshipMap.get(currType+ SPLIT_FLAG +targetType);
    	if(relationshipField == null){
    		relationshipField = resultRelationShipFile(currType, targetType);
    		relationshipMap.put(currType+ SPLIT_FLAG +targetType, relationshipField);
    	}
    	return relationshipField;
    }
    
    public Map<String,String> resultFile(String type){
        Map<String,String> map = new HashMap<String, String>();
        String jdsx = type;
        //1.创建Reader对象
        SAXReader reader = new SAXReader();
        //2.加载xml
        Document document = null;
        SAXReader saxReader = new SAXReader();
        try {

            //返回读取指定资源的输入流
            InputStream in = AnalysisXML.class.getClassLoader().getSystemResourceAsStream("file.xml");
//           writeToLocal(pathxml,in);
            document = reader.read(in);

            //3.获取根节点
            Element rootElement = document.getRootElement();
            Iterator iterator = rootElement.elementIterator();
            while (iterator.hasNext()){
                Element stu = (Element) iterator.next();
                List<Attribute> attributes = stu.attributes();
//               System.out.println("======获取属性值======");
                String s = "";
                if(stu.attribute("swr").getValue().equals(type)){
                    Iterator iterator1 = stu.elementIterator();

                    while (iterator1.hasNext()){
                        Element stuChild = (Element) iterator1.next();
                        map.put(stuChild.attribute("swr").getValue(),stuChild.attribute("alm").getValue());

                    }
                }
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
        return map;
    }

    public String resultType(String type){
        Map<String,String> map = new HashMap<String, String>();
        String jdsx = type;
        //1.创建Reader对象
        SAXReader reader = new SAXReader();
        //2.加载xml
        Document document = null;
        String s = "";
        try {

            //返回读取指定资源的输入流
            InputStream in = AnalysisXML.class.getClassLoader().getSystemResourceAsStream("file.xml");
//           writeToLocal(pathxml,in);
            document = reader.read(in);

            //3.获取根节点
            Element rootElement = document.getRootElement();
            Iterator iterator = rootElement.elementIterator();
            while (iterator.hasNext()){
                Element stu = (Element) iterator.next();
                List<Attribute> attributes = stu.attributes();
//               System.out.println("======获取属性值======");
                if(stu.attribute("swr").getValue().equals(type)){
                     for (Attribute attribute : attributes) {
                         if(attribute.getName().equals("alm")) {
                             s = attribute.getValue();
                         }
                     }
                }
//
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
        return s;
    }

    public static String resultRelationShipFile(String type1,String type2){
        String RelationShipFile = "";
        //1.创建Reader对象
        SAXReader reader = new SAXReader();
        //2.加载xml
        Document document = null;
        SAXReader saxReader = new SAXReader();
        try {

            //返回读取指定资源的输入流
            InputStream in = AnalysisXML.class.getClassLoader().getSystemResourceAsStream("RelationshipFile.xml");
//           writeToLocal(pathxml,in);
            document = reader.read(in);

            //3.获取根节点
            Element rootElement = document.getRootElement();
            Iterator iterator = rootElement.elementIterator();
            while (iterator.hasNext()){
                Element stu = (Element) iterator.next();
                List<Attribute> attributes = stu.attributes();
//               System.out.println("======获取属性值======");
                String s = "";
                if(stu.attribute("alm").getValue().equals(type1)){
                    Iterator iterator1 = stu.elementIterator();

                    while (iterator1.hasNext()){
                        Element stuChild = (Element) iterator1.next();
                        if(stuChild.attribute("type").getValue().equals(type2)){
                            RelationShipFile = stuChild.attribute("relationShipFile").getValue();
                        }
                    }
                }
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
        return RelationShipFile;
    }

    /**
     * 获取xml中type字段
     */
    public Map<String,List<String>> resultXmlType(String type1){
       Map<String,String> map = new HashMap<>();
       Map<String,List<String>> resultMap = new HashMap<>();
       Set<String> set = new HashSet();
        //1.创建Reader对象
        SAXReader reader = new SAXReader();
        //2.加载xml
        Document document = null;
        SAXReader saxReader = new SAXReader();
        try {

            //返回读取指定资源的输入流
            InputStream in = AnalysisXML.class.getClassLoader().getSystemResourceAsStream("FieldMapping.xml");
//           writeToLocal(pathxml,in);
            document = reader.read(in);

            //3.获取根节点
            Element rootElement = document.getRootElement();
            Iterator iterator = rootElement.elementIterator();
            while (iterator.hasNext()){
                Element stu = (Element) iterator.next();
                List<Attribute> attributes = stu.attributes();
//               System.out.println("======获取属性值======");
                if(stu.attribute("name").getValue().equals(type1)){
                    Iterator iterator1 = stu.elementIterator();
                    JSONObject jsonObject = new JSONObject();
                    while (iterator1.hasNext()){
                        Element stuChild = (Element) iterator1.next();
                        set.add(stuChild.attribute("type").getValue());
                        map.put(stuChild.attribute("field").getValue(),stuChild.attribute("type").getValue());
                    }
                }
            }
            if(set != null && set.size() != 0){
                for (String str : set) {
                    List l = new ArrayList();
                    for(String s : map.keySet()){
                        if(map.get(s).equals(str)){
                            l.add(s);
                        };
                    }
                    resultMap.put(str,l);
                }
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
        return resultMap;
    }



    /**
     * 获取xml中aml字段域导出字段map
     */
    public Map<String,String> resultXmlPP(String type1){
        Map<String,String> map = new HashMap<>();
        //1.创建Reader对象
        SAXReader reader = new SAXReader();
        //2.加载xml
        Document document = null;
        SAXReader saxReader = new SAXReader();
        try {

            //返回读取指定资源的输入流
            InputStream in = AnalysisXML.class.getClassLoader().getSystemResourceAsStream("FieldMapping.xml");
//           writeToLocal(pathxml,in);
            document = reader.read(in);

            //3.获取根节点
            Element rootElement = document.getRootElement();
            Iterator iterator = rootElement.elementIterator();
            while (iterator.hasNext()){
                Element stu = (Element) iterator.next();
                List<Attribute> attributes = stu.attributes();
//               System.out.println("======获取属性值======");
                if(stu.attribute("name").getValue().equals(type1)){
                    Iterator iterator1 = stu.elementIterator();
                    JSONObject jsonObject = new JSONObject();
                    while (iterator1.hasNext()){
                        Element stuChild = (Element) iterator1.next();
                        map.put(stuChild.attribute("name").getValue(),stuChild.attribute("field").getValue());
                    }
                }
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
        return map;
    }

    /**
     * 根据类型查询alm中字段 v1:测试类型 v2：属性类型 v3属性值 v4返回的属性
     */
    public List<String> resultXmlType(String type1,String type2,String type3,String type4){
        List<String> list = new ArrayList<>();
        //1.创建Reader对象
        SAXReader reader = new SAXReader();
        //2.加载xml
        Document document = null;
        SAXReader saxReader = new SAXReader();
        try {

            //返回读取指定资源的输入流
            InputStream in = AnalysisXML.class.getClassLoader().getSystemResourceAsStream("FieldMapping.xml");
//           writeToLocal(pathxml,in);
            document = reader.read(in);

            //3.获取根节点
            Element rootElement = document.getRootElement();
            Iterator iterator = rootElement.elementIterator();
            while (iterator.hasNext()){
                Element stu = (Element) iterator.next();
                List<Attribute> attributes = stu.attributes();
//               System.out.println("======获取属性值======");
                if(stu.attribute("name").getValue().equals(type1)){
                    Iterator iterator1 = stu.elementIterator();
                    JSONObject jsonObject = new JSONObject();
                    while (iterator1.hasNext()){
                        Element stuChild = (Element) iterator1.next();
                        if(type3.equals(stuChild.attribute(type2).getValue())){
                            list.add(stuChild.attribute(type4).getValue());
                        }
                    }
                }
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
        return list;
    }

    /**
     * 文档的标题
     */
    public List<String> resultXmlH(String type1){
        List<String> list = new ArrayList<>();
        //1.创建Reader对象
        SAXReader reader = new SAXReader();
        //2.加载xml
        Document document = null;
        SAXReader saxReader = new SAXReader();
        try {

            //返回读取指定资源的输入流
            InputStream in = AnalysisXML.class.getClassLoader().getSystemResourceAsStream("FieldMapping.xml");
//           writeToLocal(pathxml,in);
            document = reader.read(in);

            //3.获取根节点
            Element rootElement = document.getRootElement();
            Iterator iterator = rootElement.elementIterator();
            while (iterator.hasNext()){
                Element stu = (Element) iterator.next();
                List<Attribute> attributes = stu.attributes();
//               System.out.println("======获取属性值======");
                if(stu.attribute("name").getValue().equals(type1)){
                    Iterator iterator1 = stu.elementIterator();
                    JSONObject jsonObject = new JSONObject();
                    while (iterator1.hasNext()){
                        Element stuChild = (Element) iterator1.next();
                        list.add(stuChild.attribute("name").getValue());
                    }
                }
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
        return list;
    }


   public static void main(String[] arg){
//       Map<String,List<String>> sl = new AnalysisXML().resultXmlType("Electric System Requirement Document");
       List<String> sl = new AnalysisXML().resultXmlType("Electric System Requirement Document","type","Test Case","field");

       //       new AnalysisXML().resultCategory("Feature Function List");
       System.out.println(sl);
   }

}
