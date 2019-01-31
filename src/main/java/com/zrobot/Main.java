package com.zrobot;

import org.apache.poi.poifs.crypt.Decryptor;
import org.apache.poi.poifs.crypt.EncryptionInfo;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import java.util.concurrent.CountDownLatch;
import java.util.concurrent.ExecutorService;
import java.util.concurrent.Executors;

/**
 * @createuser gaotong
 * @createtime 2019/1/30
 * @desc
 */
public class Main {
    private static int LINE_COUNT = 1000;
    private static List<Map<String,String>> LIST_COUNT;

    public static void main(String[] args) throws Exception{
        long startTime = System.currentTimeMillis();
        InputStream is = Main.class.getClassLoader().getResourceAsStream("20190118.xlsx");

        String password = "20190118";
        POIFSFileSystem pfs = new POIFSFileSystem(is);
        is.close();
        EncryptionInfo encInfo = new EncryptionInfo(pfs);
        Decryptor decryptor = Decryptor.getInstance(encInfo);
        decryptor.verifyPassword(password);
        Workbook workbook = new XSSFWorkbook(decryptor.getDataStream(pfs));

        int threadCount = Runtime.getRuntime().availableProcessors()*2;
        CountDownLatch countDownLatch = new CountDownLatch(threadCount);
        ExecutorService pool = Executors.newFixedThreadPool(threadCount);
        Sheet sheet = workbook.getSheetAt(0);
        Task.setSheet(sheet);
        int total = sheet.getLastRowNum()-1;
        LIST_COUNT = new ArrayList<>(total);
        LINE_COUNT = total/threadCount+1;

        for(int count=0; count<threadCount; count++){
            int start = count*LINE_COUNT+1;
            int end = ((count+1)*LINE_COUNT)+1>total?total:(count+1)*LINE_COUNT+1;
            List<Map<String,String>> list = new ArrayList<>(end-start);
            pool.submit(new Task(start,end,countDownLatch,list));
            synchronized(Main.class){
                LIST_COUNT.addAll(list);
            }
        }
        countDownLatch.await();
        pool.shutdown();

        long endTime = System.currentTimeMillis();

        System.out.println(LIST_COUNT.size());
        System.out.println(endTime-startTime);

    }
}
