package com.qxtx.idea.demo;

import android.Manifest;
import android.content.pm.PackageManager;
import android.os.Build;
import android.os.Bundle;
import android.support.annotation.NonNull;
import android.support.v7.app.AppCompatActivity;
import android.util.Log;

import com.qxtx.idea.ideaexcel.poi.callback.IReadCallback;
import com.qxtx.onlytest.R;

import java.util.List;

public class MainActivity extends AppCompatActivity {

    private long durationMs = 0;

    private int requestCode = 123;

    @Override
    protected void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        setContentView(R.layout.activity_main);

        if (Build.VERSION.SDK_INT >= Build.VERSION_CODES.M &&
            checkSelfPermission(Manifest.permission.WRITE_EXTERNAL_STORAGE) != PackageManager.PERMISSION_GRANTED) {
            requestPermissions(new String[] {Manifest.permission.WRITE_EXTERNAL_STORAGE}, requestCode);
        } else {
            readExcel();
        }
    }

    @Override
    public void onRequestPermissionsResult(int requestCode, @NonNull String[] permissions, @NonNull int[] grantResults) {
        if (requestCode == this.requestCode) {
            if (grantResults[0] != PackageManager.PERMISSION_GRANTED) {
                finish();
                return ;
            }

            readExcel();
        }
    }

    private void readExcel() {
        new Thread(() -> {
            Log.e("ExcelParser", "开始计时...");
            durationMs = System.currentTimeMillis();

            String excelPath = "/sdcard/test.csv";
            ExcelHelper.readASync(excelPath, new IReadCallback() {
                @Override
                public void onRowRead(int rowIndex, @NonNull List<String> rowInfo) {
                    //打印log会极大地增加流程耗时
                    String row = "";
                    for (int i = 0; i < rowInfo.size(); i++) {
                        if (i == 0) {
                            row = rowInfo.get(i);
                            continue;
                        }
                        row += "###" + rowInfo.get(i);
                    }
                    Log.e("ExcelParser", "行：" + rowIndex + ", 内容：[" + row + "]====数量" + rowInfo.size());
                }

                @Override
                public void onFinished() {
                    Log.i("TAG", "表格读取完成");
                }
            });

            Log.e("ExcelParser", "耗时：" + (System.currentTimeMillis() - durationMs) + "ms.");
        }).start();
    }
}
