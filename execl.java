package com.fcjr;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.math.BigDecimal;
import java.net.URL;
import java.util.ArrayList;
import java.util.Date;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.Map.Entry;

import org.springframework.stereotype.Component;

import jxl.Workbook;
import jxl.write.Label;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;
@Component
public class execl {
	
	public static void main(String[] args) {
	        //数据准备
	        List<Map<String, Object>> list = new ArrayList<>();
	            Map<String, Object> map = new LinkedHashMap<>();
	            map.put("开始造影","2019/3/21 10:10:22");
	            map.put("结束造影","2019/3/21 11:20:21");
	            map.put("导丝通过","2019/3/22 14:22:01");
	            list.add(map);
	        //存放路径
	        String path;
	        try {
	            //获取项目的路径
	            path = Class.class.getClass().getResource("/").getPath();
	            //路径转换下格式
	            path = path.replaceAll("classes", "excel").concat("case").concat(".xls");
	            path=path.substring(1);
	            System.out.println(path);
	            //写入到excel
	            writeExcel(list, "case.xls");
	        } catch (Exception e) {
	            e.printStackTrace();
	        }
	    }
	
	public static void writeExcel(List<Map<String, Object>> list, String path) {
        try {
            // Excel底部的表名
            String sheetn = "case";
            // 用JXL向新建的文件中添加内容
            File myFilePath = new File(path);
            if (!myFilePath.exists())
                myFilePath.createNewFile();
            OutputStream outputstream = new FileOutputStream(path);
            WritableWorkbook writableworkbook = Workbook.createWorkbook(outputstream);
            jxl.write.WritableSheet writesheet = writableworkbook.createSheet(sheetn, 1);
            // 设置标题
            if (list.size() > 0) {
                int j = 0;
                for (Entry<String, Object> entry : list.get(0).entrySet()) {
                    String title = entry.getKey();
                    writesheet.addCell(new Label(j, 0, title));
                    j++;
                }
            }
            // 内容添加
            for (int i = 1; i <= list.size(); i++) {
                int j = 0;
                for (Entry<String, Object> entry : list.get(i - 1).entrySet()) {
                    Object o = entry.getValue();
                    if (o instanceof Double) {
                        writesheet.addCell(new jxl.write.Number(j, i, (Double) entry.getValue()));
                    } else if (o instanceof Integer) {
                        writesheet.addCell(new jxl.write.Number(j, i, (Integer) entry.getValue()));
                    } else if (o instanceof Float) {
                        writesheet.addCell(new jxl.write.Number(j, i, (Float) entry.getValue()));
                    } else if (o instanceof Float) {
                        writesheet.addCell(new jxl.write.DateTime(j,i,(Date) entry.getValue()));
                    } else if (o instanceof BigDecimal) {
                        writesheet.addCell(new jxl.write.Number(j, i, ((BigDecimal) entry
                                .getValue()).doubleValue()));
                    } else if (o instanceof Long) {
                        writesheet.addCell(new jxl.write.Number(j, i, ((Long) entry.getValue())
                                .doubleValue()));
                    } else {
                        writesheet.addCell(new Label(j, i, (String) entry.getValue()));
                    }
                    j++;
                }
            }
            writableworkbook.write();
            writableworkbook.close();
        } catch (WriteException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
