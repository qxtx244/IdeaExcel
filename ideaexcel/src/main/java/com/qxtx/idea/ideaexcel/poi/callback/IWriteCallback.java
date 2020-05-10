package com.qxtx.idea.ideaexcel.poi.callback;

/**
 * Created in 2020/4/29 17:25
 *
 * @author QXTX-WORK
 * <p>
 * Description
 */
public interface IWriteCallback {

    void onWriteStart();

    void onWriteFinished();

    void onWriteError();
}
