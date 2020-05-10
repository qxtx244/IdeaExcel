package com.qxtx.idea.demo;

import android.support.annotation.NonNull;
import android.util.Log;

import com.qxtx.idea.ideaexcel.poi.bean.RowBean;
import com.qxtx.idea.ideaexcel.poi.callback.IReadCallback;
import com.qxtx.idea.ideaexcel.poi.callback.IWriteCallback;
import com.qxtx.idea.ideaexcel.poi.parser.ExcelReader;
import com.qxtx.idea.ideaexcel.poi.parser.ExcelWriter;

import java.io.File;
import java.util.List;
import java.util.concurrent.ExecutorService;
import java.util.concurrent.Executors;

/**
 * Created in 2020/5/6 11:31
 *
 * @author QXTX-WORK
 * <p>
 * Description 本地表格文件处理 管理类，集合常用api
 */
public class ExcelHelper {

    private final static String TAG = "ExcelHelper";

    /** 每一次读取，使用一个新线程 */
    private static ExecutorService mPool = Executors.newCachedThreadPool();

    /**
     * 读取excel文件，回调读取事件
     * @param excelPath 本地表格文件的绝对路径
     * @param callback 读取每一有效行数据的回调
     */
    public static void readASync(@NonNull String excelPath, @NonNull IReadCallback callback) {
        mPool.execute(() -> {
            final long durationMs = System.currentTimeMillis();

            File excelFile = new File(excelPath);
            if (!excelFile.exists()) {
                Log.i(TAG, "文件不存在" + excelFile.getPath());
                return;
            }

            ExcelReader.getInstance().read(excelPath, callback);
            Log.i(TAG, "读取表格耗时[" + (System.currentTimeMillis() - durationMs) + "]ms.");

            callback.onFinished();
        });
    }

    /**
     * 导出xlsx表格
     * @param excelPath 导出xlsx表格的目标路径
     * @param rowBeanList 表格行数据对象列表
     * @deprecated 仅用作测试，不可使用
     */
    @Deprecated
    public static void createXlsxASync(@NonNull String excelPath, @NonNull List<RowBean> rowBeanList) {
        mPool.execute(() -> {
            long durationMs = System.currentTimeMillis();

            try {
                ExcelWriter excelWriter = ExcelWriter.getInstance();
                excelWriter.writeXlsx(rowBeanList, excelPath, new IWriteCallback() {
                    @Override
                    public void onWriteStart() {
                        Log.e(getClass().getSimpleName(), "开始导入.");
                    }

                    @Override
                    public void onWriteFinished() {
                        Log.i(getClass().getSimpleName(), "导出excel耗时：" + (System.currentTimeMillis() - durationMs) + "ms.");
                    }

                    @Override
                    public void onWriteError() {
                        Log.e(getClass().getSimpleName(), "导入excel异常");
                    }
                });
            } catch (Exception e) {
                Log.e(TAG, "导出excel异常：" + e.getLocalizedMessage());
                e.printStackTrace();
            }

            Log.i(TAG, "导出excel耗时[" + (System.currentTimeMillis() - durationMs) + "]ms.");
        });
    }

    /**
     * 请求停止所有线程，但可能不会立即完成
     */
    public static void release() {
        if (!mPool.isShutdown()) {
            mPool.shutdown();
        }
        mPool = null;
    }
}
