package com.qxtx.idea.ideaexcel.poi.parser;

import android.support.annotation.NonNull;
import android.util.Log;

import com.qxtx.idea.ideaexcel.poi.bean.RowBean;
import com.qxtx.idea.ideaexcel.poi.callback.IWriteCallback;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;

/**
 * Created in 2020/4/29 13:32
 *
 * @author QXTX-WORK
 * <p>
 * Description
 */
public class ExcelWriter {

    private volatile static ExcelWriter instance;

    public static ExcelWriter getInstance() {
        if (instance == null) {
            synchronized (ExcelWriter.class) {
                if (instance == null) {
                    instance = new ExcelWriter();
                }
            }
        }
        return instance;
    }

    private ExcelWriter() { }

    /**
     * 导出数据到xlsx表格
     * @deprecated 仅作为测试用
     */
    @Deprecated
    public void writeXlsx(List<RowBean> rowList, @NonNull String excelPath, @NonNull IWriteCallback callback) {
        File file = new File(excelPath);
        if (file.isDirectory()) {
            Log.e("ExcelParser", "路径为目录，不支持");
            return ;
        }
        file.mkdirs();

        callback.onWriteStart();

        if (file.exists()) {
            file.delete();
        }

        SXSSFWorkbook swb = new SXSSFWorkbook(new XSSFWorkbook(), 1000);
        SXSSFSheet sheet = (SXSSFSheet) swb.createSheet("info");

        for (int i = 0; i < rowList.size(); i++) {
            RowBean bean = rowList.get(i);
            Row dataRow = sheet.createRow(i);
            dataRow.createCell(0).setCellValue(bean.getName());
            dataRow.createCell(1).setCellValue(bean.getId());
            dataRow.createCell(2).setCellValue(bean.getCensusType());
            dataRow.createCell(3).setCellValue(bean.getAddress());
        }

        FileOutputStream fos = null;
        try {
            fos = new FileOutputStream(file);
            swb.write(fos);
            swb.close();
            fos.close();
        } catch (IOException e) {
            e.printStackTrace();
            callback.onWriteError();
            return ;
        }

        callback.onWriteFinished();
    }
}
