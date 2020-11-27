package com.lv;

import java.util.List;

public class Result<T> {


    List<String> ColumnList;
    List<T> productList;

    public List<String> getColumnList() {
        return ColumnList;
    }

    public void setColumnList(List<String> columnList) {
        ColumnList = columnList;
    }

    public List<T> getProductList() {
        return productList;
    }

    public void setProductList(List<T> productList) {
        this.productList = productList;
    }

    @Override
    public String toString() {
        return "Result{" +
                "ColumnList=" + ColumnList +
                ", productList=" + productList +
                '}';
    }
}
