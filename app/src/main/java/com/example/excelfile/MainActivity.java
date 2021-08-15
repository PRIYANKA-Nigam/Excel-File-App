package com.example.excelfile;

import androidx.appcompat.app.AppCompatActivity;
import androidx.core.app.ActivityCompat;

import android.Manifest;
import android.content.Intent;
import android.content.pm.PackageManager;
import android.os.Bundle;
import android.os.Environment;
import android.provider.CalendarContract;
import android.view.View;
import android.widget.Button;
import android.widget.EditText;
import android.widget.TextClock;
import android.widget.TextView;
import android.widget.Toast;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Calendar;

public class MainActivity extends AppCompatActivity {
   EditText e;
   private File filePath=new File(Environment.getExternalStorageDirectory() + "/Demos.xls");
EditText e1,e2,e3;Button b;@Override
    protected void onCreate(Bundle savedInstanceState) { super.onCreate(savedInstanceState);setContentView(R.layout.activity_main);
        ActivityCompat.requestPermissions(this,new String[]{Manifest.permission.WRITE_EXTERNAL_STORAGE
                ,Manifest.permission.READ_EXTERNAL_STORAGE}, PackageManager.PERMISSION_GRANTED);
        e=(EditText)findViewById(R.id.editTextTextPersonName);e1=(EditText)findViewById(R.id.editTextTextPersonName2);
        e2=(EditText)findViewById(R.id.editTextTextPersonName3);e3=(EditText)findViewById(R.id.editTextTextMultiLine);
        b=(Button)findViewById(R.id.button2);b.setOnClickListener(new View.OnClickListener() {@Override
            public void onClick(View v) { if (!e1.getText().toString().isEmpty()&&!e2.getText().toString().isEmpty()&&!e3.getText().toString().isEmpty()){
                    Intent intent=new Intent(Intent.ACTION_INSERT);
                    intent.setData(CalendarContract.Events.CONTENT_URI);
                    intent.putExtra(CalendarContract.Events.TITLE,e1.getText().toString());
                    intent.putExtra(CalendarContract.Events.EVENT_LOCATION,e2.getText().toString());
                    intent.putExtra(CalendarContract.Events.DESCRIPTION,e3.getText().toString());
                    intent.putExtra(CalendarContract.Events.ALL_DAY,"true");
                    intent.putExtra(Intent.EXTRA_EMAIL,"priyankanigam25041999@gmail.com,nigamvivek509@gmail.com");
                    if (intent.resolveActivity(getPackageManager())!=null){ startActivity(intent);
                    }else { Toast.makeText(MainActivity.this,"No Apps can perform this action",Toast.LENGTH_SHORT).show(); }
            }else { Toast.makeText(MainActivity.this,"Enter all the Fields",Toast.LENGTH_SHORT).show(); } }}); }
            public void CreateExcel(View view) throws IOException { HSSFWorkbook  hssfWorkbook=new HSSFWorkbook();
        HSSFSheet hssfSheet=hssfWorkbook.createSheet("Custom Sheet");
        Toast.makeText(this,"Cell record created",Toast.LENGTH_SHORT).show();
        HSSFRow hssfRow =hssfSheet.createRow(0);HSSFCell  hssfCell =hssfRow.createCell(0);
                hssfCell.setCellValue(e.getText().toString());
        try { if (!filePath.exists()){ filePath.createNewFile(); }
            FileOutputStream fileOutputStream=new FileOutputStream(filePath);hssfWorkbook.write(fileOutputStream);
            if (fileOutputStream!=null){ fileOutputStream.flush();fileOutputStream.close(); }
        }catch (Exception e){ e.printStackTrace(); } }}