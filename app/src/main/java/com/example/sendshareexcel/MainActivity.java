package com.example.sendshareexcel;

import android.Manifest;
import android.content.Intent;
import android.content.pm.PackageManager;
import android.net.Uri;
import android.os.Bundle;
import android.os.Environment;
import android.view.View;

import androidx.appcompat.app.AppCompatActivity;
import androidx.core.app.ActivityCompat;
import androidx.core.content.FileProvider;

import com.google.android.material.button.MaterialButton;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

import java.io.File;
import java.io.FileOutputStream;

public class MainActivity extends AppCompatActivity {

    private MaterialButton sendBtn;

    @Override
    protected void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        setContentView(R.layout.activity_main);

        ActivityCompat.requestPermissions(this, new String[]{Manifest.permission.WRITE_EXTERNAL_STORAGE, Manifest.permission.READ_EXTERNAL_STORAGE}, PackageManager.PERMISSION_GRANTED);

        sendBtn = findViewById(R.id.send_button);
        sendBtn.setOnClickListener(new View.OnClickListener() {
            @Override
            public void onClick(View view) {
                HSSFWorkbook hssfWorkbook = new HSSFWorkbook();
                HSSFSheet hssfSheet = hssfWorkbook.createSheet();

                HSSFRow hssfRow = hssfSheet.createRow(0);
                HSSFCell hssfCell = hssfRow.createCell(0);

                hssfCell.setCellValue("TEST");

                String filePath2 = Environment.getExternalStoragePublicDirectory(Environment.DIRECTORY_DOCUMENTS) +  "/Demo.xls";

                File filepath = new File(filePath2);

                try {
                    if (!filepath.exists()) {
                        filepath.createNewFile();
                    }
                    FileOutputStream fileOutputStream = new FileOutputStream(filepath);

                    hssfWorkbook.write(fileOutputStream);

                    if (fileOutputStream != null) {
                        fileOutputStream.flush();
                        fileOutputStream.close();
                    }
                    // creation of the file's uri to send
                    Uri fileURI = FileProvider.getUriForFile(MainActivity.this, getApplicationContext().getPackageName() + ".provider", filepath);

                    Intent sendIntent = new Intent();
                    sendIntent.setAction(Intent.ACTION_SEND);
                    sendIntent.putExtra(Intent.EXTRA_TEXT, "THiS iS mY TeSt tO sEnD.");
                    sendIntent.setDataAndType(fileURI, "text/xls");
                    sendIntent.putExtra(Intent.EXTRA_STREAM, fileURI);

                    Intent shareIntent = Intent.createChooser(sendIntent, "tEsT!");
                    startActivity(shareIntent);
                } catch (Exception e) {
                    e.printStackTrace();
                }


                // first create file object for file placed at location
                // specified by filepath
//                String filePath = Environment.getExternalStoragePublicDirectory(Environment.DIRECTORY_DOCUMENTS) + "/TEST_CSV.csv";
//                File file = new File(filePath);
//                try {
//                    // create FileWriter object with file as parameter
//                    FileWriter outputfile = new FileWriter(file);
//
//                    // create CSVWriter object filewriter object as parameter
//                    CSVWriter writer = new CSVWriter(outputfile);
//
//                    // adding header to csv
//                    String[] header = { "Name", "Class", "Marks" };
//                    writer.writeNext(header);
//
//                    // add data to csv
//                    String[] data1 = { "Aman", "10", "620" };
//                    writer.writeNext(data1);
//                    String[] data2 = { "Suraj", "10", "630" };
//                    writer.writeNext(data2);
//                    String[] data3 = { "Edo", "09", "690" };
//                    writer.writeNext(data3);
//
//                    // closing writer connection
//                    writer.close();
//
//                    // creation of the file's uri to send
//                    Uri fileURI = FileProvider.getUriForFile(MainActivity.this, getApplicationContext().getPackageName() + ".provider", file);
//
//                    Intent sendIntent = new Intent();
//                    sendIntent.setAction(Intent.ACTION_SEND);
//                    sendIntent.putExtra(Intent.EXTRA_TEXT, "THiS iS mY TeSt tO sEnD.");
//                    sendIntent.setDataAndType(fileURI, "text/csv");
//                    sendIntent.putExtra(Intent.EXTRA_STREAM, fileURI);
//
//                    Intent shareIntent = Intent.createChooser(sendIntent, "tEsT!");
//                    startActivity(shareIntent);
//
//                }
//                catch (IOException e) {
//                    e.printStackTrace();
//                }
            }
        });
    }
}