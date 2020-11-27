package com.lv;

import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.junit.Test;

import java.io.File;
import java.io.FileOutputStream;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.ResultSet;
import java.sql.Statement;
import java.util.ArrayList;
import java.util.List;
// 0.1的版本 基本上都写死了
public class HSSFWorkbookTest {
    public static void main(String[] args) throws Exception {
        HSSFWorkbook wb = new HSSFWorkbook();
        File file = new File("D:\\360MoveData\\Users\\吕亚东\\Desktop\\test.xls");
        FileOutputStream outputStream = new FileOutputStream(file);
        HSSFSheet sheet = wb.createSheet("第一个工作页");

        HSSFRow row = sheet.createRow(0);
        JDBC();
        Result<Product> result = JDBC();
        List<String> columnList = result.getColumnList();
        //列数
        int size = columnList.size();
        //数据集合
        List<Product> productList = result.getProductList();
        for (int i = 0; i < result.ColumnList.size(); i++) {
            row.createCell(i).setCellValue(result.ColumnList.get(i));

        }
/*        for (int i = 1; i <= result.productList.size(); i++) {
            HSSFRow newRow = sheet.createRow(i);
            newRow.createCell(0).setCellValue(productList.get(i-1).getP_id());
            newRow.createCell(1).setCellValue(productList.get(i-1).getpName());
            newRow.createCell(2).setCellValue(productList.get(i-1).getPrice());
            newRow.createCell(3).setCellValue(productList.get(i-1).getpDesc());
            newRow.createCell(4).setCellValue(productList.get(i-1).getpColor());
            newRow.createCell(5).setCellValue(productList.get(i-1).getpSpeci());
            newRow.createCell(6).setCellValue(productList.get(i-1).getC_id());
            newRow.createCell(7).setCellValue(productList.get(i-1).getStore());
            newRow.createCell(8).setCellValue(productList.get(i-1).getpImg());

        }*/
        wb.write(outputStream);
        outputStream.close();
    }


    @Test
    static Result<Product> JDBC() throws Exception {
        Class.forName("com.mysql.jdbc.Driver");
        Connection connection = DriverManager.
                getConnection("jdbc:mysql://localhost:3306/myproject", "root", "root");

        Statement statement = connection.createStatement();
        String sql = "select * from t_product";

        ResultSet resultSet = statement.executeQuery(sql);

        //列的名称
        List<String> Columnlist = getColumnName(resultSet);

        List<Product> productList = new ArrayList<>();
        while (resultSet.next()) {
            Product product = new Product();
            //这个可以获取当前列的数据类型
            System.out.println(resultSet.getMetaData().getColumnClassName(3));

            product.setP_id(resultSet.getInt("p_id"));
            product.setpName(resultSet.getString("pName"));
            product.setPrice(resultSet.getFloat("price"));
            product.setpDesc(resultSet.getString("pDesc"));
            product.setpColor(resultSet.getString("pColor"));
            product.setpSpeci(resultSet.getString("pSpeci"));
            product.setC_id(resultSet.getInt("c_id"));
            product.setStore(resultSet.getInt("store"));
            product.setpImg(resultSet.getString("pImg"));
            productList.add(product);

        }
        resultSet.close();
        statement.close();
        connection.close();
        Result<Product> result = new Result<>();
        result.setColumnList(Columnlist);
        result.setProductList(productList);
        return result;

    }

    //返回当前表中所有的列名
    static List<String> getColumnName(ResultSet resultSet) throws Exception {
        List<String> Columnlist = new ArrayList<>();
        ArrayList<Object> ClassList = new ArrayList<>();
        int count = resultSet.getMetaData().getColumnCount();
        resultSet.getMetaData().getColumnType(1);
        for (int i = 1; i <= count; i++) {
            Columnlist.add(resultSet.getMetaData().getColumnName(i));
            ClassList.add(resultSet.getMetaData().getColumnClassName(i));
        }
        System.out.println(ClassList);
        return Columnlist;
    }



}
