package com.qxtx.idea.ideaexcel.poi.callback;

import android.support.annotation.NonNull;

import java.util.List;

/**
 * Created in 2020/4/26 16:57
 *
 * @author QXTX-WORK
 * <p>
 * Description 读取表格数据的回调
 */
public interface IReadCallback {

    /**
     * 读取到一行表格数据
     * @param rowIndex 读取的行数
     * @param row 表格中一行数据内容
     */
    void onRowRead(int rowIndex, @NonNull List<String> row);

    /** 表格读取结束 */
    void onFinished();
}
