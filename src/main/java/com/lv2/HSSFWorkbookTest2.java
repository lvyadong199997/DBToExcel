package com.lv2;

import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.junit.Test;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.ResultSet;
import java.sql.Statement;
import java.util.*;

// 0.2的版本 可以从数据库中取到任意类型的数据   DB -->>> Excel
public class HSSFWorkbookTest2 {

    public static void main(String[] args) throws Exception {
        HSSFWorkbookTest2 hssf = new HSSFWorkbookTest2();

        //todo 读取外部配置文件 这个地方写死了
        String path = "D:\\java代码\\mywork\\src\\main\\resources\\param.properties";
        HashMap<String, String> paramMap=paramMap = hssf.ReadPropertiesFromFile(path);
        String out = paramMap.get("out");

        //创建一个 HSSFWorkbook 对象 使用它区操作Excel表
        HSSFWorkbook wb = new HSSFWorkbook();

        //输出路径
        File file = new File(out);
        FileOutputStream outputStream = new FileOutputStream(file);

        HSSFSheet sheet = wb.createSheet("第一个工作页");

        //连接数据库 返回数据库中的所有的列名称 和 对应的数据类型 并且录入数据
        Map<String, List<String>> map = hssf.getConnectOfMysql(sheet, paramMap);
        //创建第一行
        hssf.createFirstRow(map.get("columnNameList"), sheet);
        //写出数据
        wb.write(outputStream);
        //资源关闭
        outputStream.close();
    }

    //读取配置文件
    public HashMap<String, String> ReadPropertiesFromFile(String file) throws Exception {
        Properties properties = new Properties();

        properties.load(new FileInputStream(file));
        //读取数据
        String db = properties.get("DB").toString();
        String username = properties.get("username").toString();
        String password = properties.get("password").toString();
        String table = properties.get("table").toString();
        String out = properties.get("out").toString();
        //封装数据
        HashMap<String, String> paramMap = new HashMap<>();
        paramMap.put("db", db);
        paramMap.put("username", username);
        paramMap.put("password", password);
        paramMap.put("table", table);
        paramMap.put("out",out);
        return paramMap;
    }


    //创建第一行 也就是每一行的标题
    public void createFirstRow(List<String> columnNameList, HSSFSheet sheet) {
        HSSFRow row = sheet.createRow(0);
        for (int i = 1; i <= columnNameList.size(); i++) {
            row.createCell(i - 1).setCellValue(columnNameList.get(i - 1));
        }
    }

    /**
     * sheet 当前工作页
     * paramMap 一些操作的参数
     * 连接数据库 返回数据库中的所有的列名称 和 对应的数据类型 并且录入数据
     */
    public Map<String, List<String>> getConnectOfMysql(HSSFSheet sheet, HashMap<String, String> paramMap) throws Exception {
        //todo 动态更改的数据有 数据库 用户名 密码  表 记住 已经解决
        Class.forName("com.mysql.jdbc.Driver");

        Connection connection = DriverManager.
                getConnection(paramMap.get("db"), paramMap.get("username"), paramMap.get("password"));

        Statement statement = connection.createStatement();
        String sql = "select * from " + paramMap.get("table");

        ResultSet resultSet = statement.executeQuery(sql);

        Map<String, List<String>> map = getColumnName(resultSet);

        List<String> columnNameList = map.get("columnNameList");
        //每个列的数据类型
        List<String> columnTypeList = map.get("columnTypeList");
        //录入数据 开始

        //s 来控制列(每一次都是新的一列)
        long before = System.currentTimeMillis();
        int s = 1;
        while (resultSet.next()) {
            //新创建一列
            HSSFRow newRow = sheet.createRow(s);

            //此时的Result对象 就相当于一个Product
            for (int i = 1; i <= columnNameList.size(); i++) {
                // 通过反射动态获取表中列的数据类型
                // System.out.println(resultSet.getObject(i,Class.forName(columnTypeList.get(i-1))));
                Object o = resultSet.getObject(i, Class.forName(columnTypeList.get(i - 1)));
                if (o != null) {
                    //o 是可能为空的 注意
                    newRow.createCell(i - 1).setCellValue(resultSet.getObject(i, Class.forName(columnTypeList.get(i - 1))).toString());
                } else {
                    newRow.createCell(i - 1).setCellValue("空");
                }

            }

            s++;

        }


        //结束
        resultSet.close();
        statement.close();
        connection.close();

        return map;
    }

    /**
     * 获取数据库中的所有的列名称和对应的数据类型
     */
    public Map<String, List<String>> getColumnName(ResultSet resultSet) throws Exception {
        //封装结果集
        Map<String, List<String>> map = new HashMap<>();
        List<String> Columnlist = new ArrayList<>();
        List<String> ClassList = new ArrayList<>();
        int count = resultSet.getMetaData().getColumnCount();
        resultSet.getMetaData().getColumnType(1);
        for (int i = 1; i <= count; i++) {
            //获取列的名称
            Columnlist.add(resultSet.getMetaData().getColumnName(i));
            //获取列的数据类型 后期通过反射来获取列的数据和名称
            ClassList.add(resultSet.getMetaData().getColumnClassName(i));
        }
        map.put("columnNameList", Columnlist);
        map.put("columnTypeList", ClassList);
        return map;
    }
}
