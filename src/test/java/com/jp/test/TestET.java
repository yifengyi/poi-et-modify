package com.jp.test;

import com.jg.poiet.XSSFTemplate;

import java.io.IOException;
import java.util.HashMap;
import java.util.Map;

/**
 * 1.0v created by admin on 2021-1-26
 */
public class TestET {
  public static void main(String[] args) {
    //！！以GBK输出
    String path = "C:/Users/admin/Desktop/简易检测/模板/";
    String src = "到货入库简易检验验收单-馈线-模板.xlsx";
    // String src = "到货入库简易检验验收单-馈线-模板 - 副本.xlsx";
    String  desc = "到货入库简易检验验收单.xlsx";
    XSSFTemplate tmp = XSSFTemplate.compile(path + src);

    Map<String, Object> map = new HashMap<>();



    map.put("unitName","□郑州\r\n ■ 三门峡");
    map.put("location","河南省郑州市");
    map.put("person","江华");
    map.put("vendorName","华为");
    map.put("versionNum","1.0v");
    map.put("poNum","PO3251000002021000001");
    map.put("date","2021年1月26日");


    try {
      tmp.render(map);
      tmp.writeToFile(path+desc);
    } catch (IOException e) {
      e.printStackTrace();
    }
  }
}
