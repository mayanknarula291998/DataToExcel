package com.icss.datatoexcel;

import androidx.appcompat.app.AppCompatActivity;
import androidx.core.app.ActivityCompat;

import android.Manifest;
import android.content.pm.PackageManager;
import android.os.Bundle;
import android.os.Environment;
import android.util.Log;
import android.view.View;
import android.widget.EditText;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

import java.io.File;
import java.io.FileOutputStream;

public class MainActivity extends AppCompatActivity {

    private EditText editTextExcel;
    private File filePath = new File(Environment.getExternalStorageDirectory() + "/Demo.xls");

    @Override
    protected void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        setContentView(R.layout.activity_main);

        ActivityCompat.requestPermissions(this, new String[]{Manifest.permission.WRITE_EXTERNAL_STORAGE,
                Manifest.permission.READ_EXTERNAL_STORAGE}, PackageManager.PERMISSION_GRANTED);

        editTextExcel = findViewById(R.id.editText);
    }

    public void buttonCreateExcel(View view) {
        Log.d("====", "DONE");
        HSSFWorkbook hssfWorkbook = new HSSFWorkbook();
        HSSFSheet hssfSheet = hssfWorkbook.createSheet("Custom Sheet");

        for (int i = 0; i < 5; i++) {
            HSSFRow hssfRow = hssfSheet.createRow(i);

            HSSFCell hssfCell1 = hssfRow.createCell(0);
            HSSFCell hssfCell2 = hssfRow.createCell(1);
            HSSFCell hssfCell3 = hssfRow.createCell(2);
            HSSFCell hssfCell4 = hssfRow.createCell(3);

            hssfCell1.setCellValue(editTextExcel.getText().toString() + "ONE-"+i);
            hssfCell2.setCellValue(editTextExcel.getText().toString() + "TWO-"+i);
            hssfCell3.setCellValue(editTextExcel.getText().toString() + "THREE-"+i);
            hssfCell4.setCellValue(editTextExcel.getText().toString() + "FOUR-"+i);

        }

        try {
            if (!filePath.exists()) {
                filePath.createNewFile();
            }

            FileOutputStream fileOutputStream = new FileOutputStream(filePath);
            hssfWorkbook.write(fileOutputStream);

            if (fileOutputStream != null) {
                fileOutputStream.flush();
                fileOutputStream.close();
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
    }



}