package com.zrobot;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.concurrent.CountDownLatch;

/**
 * @createuser gaotong
 * @createtime 2019/1/30
 * @desc
 */
public class Task implements Runnable {
    private static Sheet sheet;
    private List<Map<String,String>> list;
    private CountDownLatch countDownLatch;
    private int start;
    private int end;

    public static void setSheet(Sheet s){
        sheet = s;
    }

    public Task(int start, int end, CountDownLatch countDownLatch,List<Map<String,String>> list){
        this.start = start;
        this.end = end;
        this.countDownLatch = countDownLatch;
        this.list = list;
    }

    @Override
    public void run(){
        for(int index=start; start<end; index++){
            Row row = sheet.getRow(index);

            Map<String,String> map = new HashMap<>();
            map.put("date",row.getCell(0).getStringCellValue());
            map.put("card",row.getCell(1).getStringCellValue());
            list.add(map);
        }
        countDownLatch.countDown();
        System.out.println(end);
    }
}
