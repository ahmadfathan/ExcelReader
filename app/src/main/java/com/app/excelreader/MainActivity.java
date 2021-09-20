package com.app.excelreader;

import androidx.activity.result.ActivityResult;
import androidx.activity.result.ActivityResultCallback;
import androidx.activity.result.ActivityResultLauncher;
import androidx.activity.result.contract.ActivityResultContracts;
import androidx.annotation.NonNull;
import androidx.appcompat.app.AppCompatActivity;

import android.Manifest;
import android.app.Activity;
import android.app.AlertDialog;
import android.content.ContentValues;
import android.content.Intent;
import android.graphics.Bitmap;
import android.graphics.Canvas;
import android.graphics.Color;
import android.graphics.drawable.Drawable;
import android.net.Uri;
import android.os.Build;
import android.os.Bundle;
import android.os.Environment;
import android.provider.MediaStore;
import android.view.LayoutInflater;
import android.view.View;
import android.widget.EditText;
import android.widget.Toast;

import com.app.excelreader.databinding.ActivityMainBinding;
import com.github.mikephil.charting.data.Entry;
import com.github.mikephil.charting.data.LineData;
import com.github.mikephil.charting.data.LineDataSet;
import com.github.mikephil.charting.utils.ColorTemplate;
import com.karumi.dexter.Dexter;
import com.karumi.dexter.MultiplePermissionsReport;
import com.karumi.dexter.PermissionToken;
import com.karumi.dexter.listener.PermissionRequest;
import com.karumi.dexter.listener.multi.MultiplePermissionsListener;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.List;
import java.util.Random;

public class MainActivity extends AppCompatActivity {

    private LineData lineData;
    private ActivityMainBinding binding;
    private Uri screenshotUri;
    private XSSFWorkbook wb;

    private List<String> labels = new ArrayList<>();
    private List<List<Entry>> entries = new ArrayList<>();
    private int currentChartIndex = 0;

    @Override
    protected void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);

        binding = ActivityMainBinding.inflate(getLayoutInflater());
        View view = binding.getRoot();
        setContentView(view);

        lineData = new LineData();

        binding.ibOpen.setOnClickListener(view1 -> {
            Intent chooseFileIntent = new Intent(Intent.ACTION_GET_CONTENT);
            chooseFileIntent.setType("*/*");
            // Only return URIs that can be opened with ContentResolver
            chooseFileIntent.addCategory(Intent.CATEGORY_OPENABLE);

            chooseFileIntent = Intent.createChooser(chooseFileIntent, "Choose a file");

            chooseFileActivityResultLauncher.launch(chooseFileIntent);
        });

        binding.ibScreenshot.setOnClickListener(view12 -> {
            if(binding.lineChart.getLineData() != null){
                try {
                    Toast.makeText(MainActivity.this, "Taking a Screen Capture ..", Toast.LENGTH_SHORT).show();
                    takeScreenshot();
                } catch (IOException e) {
                    Toast.makeText(MainActivity.this, "Exception: " + e.getMessage(), Toast.LENGTH_SHORT).show();
                    e.printStackTrace();
                }
            }else{
                Toast.makeText(MainActivity.this, "No Chart Data Available", Toast.LENGTH_SHORT).show();
            }

        });

        binding.ibOpenScrenshot.setOnClickListener(view13 -> {
            if(screenshotUri != null){
                openScreenshot();
            }else{
                Toast.makeText(MainActivity.this, "No Screenshot Available", Toast.LENGTH_SHORT).show();
            }
        });

        binding.ibEditName.setOnClickListener(view16 -> {
            if(wb != null){
                AlertDialog.Builder adb = new AlertDialog.Builder(MainActivity.this);
                adb.setTitle("Save Workbook As");
                View v = LayoutInflater.from(MainActivity.this).inflate(R.layout.dialog_rename, null);
                adb.setView(v);

                EditText etFilename = v.findViewById(R.id.et_filename);

                adb.setNegativeButton("Cancel", (dialogInterface, i) -> {

                });

                adb.setPositiveButton("Rename", (dialogInterface, i) -> {
                    if(!etFilename.getText().toString().isEmpty()){
                        try{
                            renameExcelFile(etFilename.getText().toString());
                        }catch (Exception e){
                            Toast.makeText(MainActivity.this, "Error: " + e.getMessage(), Toast.LENGTH_SHORT).show();
                        }
                    }
                });

                adb.show();
            }else{
                Toast.makeText(MainActivity.this, "No Workbook Opened", Toast.LENGTH_SHORT).show();
            }
        });

        binding.ibNext.setOnClickListener(view14 -> {
            if(currentChartIndex < entries.size() - 1){
                currentChartIndex++;
                drawChart(currentChartIndex);
            }
        });

        binding.ibBack.setOnClickListener(view15 -> {
            if(currentChartIndex > 0){
                currentChartIndex--;
                drawChart(currentChartIndex);
            }
        });

        if (Build.VERSION.SDK_INT >= Build.VERSION_CODES.R) {
            Dexter.withContext(this)
                    .withPermissions(
                            Manifest.permission.MANAGE_EXTERNAL_STORAGE,
                            Manifest.permission.WRITE_EXTERNAL_STORAGE
                    ).withListener(new MultiplePermissionsListener() {
                @Override
                public void onPermissionsChecked(MultiplePermissionsReport report) {
                    /* ... */
                }
                @Override
                public void onPermissionRationaleShouldBeShown(List<PermissionRequest> permissions, PermissionToken token) {
                    /* ... */
                }

            }).check();
        }else{
            Dexter.withContext(this)
                    .withPermissions(
                            Manifest.permission.WRITE_EXTERNAL_STORAGE
                    ).withListener(new MultiplePermissionsListener() {
                @Override
                public void onPermissionsChecked(MultiplePermissionsReport report) {
                    /* ... */
                }
                @Override
                public void onPermissionRationaleShouldBeShown(List<PermissionRequest> permissions, PermissionToken token) {
                    /* ... */
                }

            }).check();
        }

    }

    private void drawChart(int i){
        lineData.clearValues();

        LineDataSet lineDataSet = new LineDataSet(entries.get(i), labels.get(i));

        Random obj = new Random();
        int rand_num = obj.nextInt(0xffffff + 1);

        lineDataSet.setColors(ColorTemplate.rgb(String.format("#%06x", rand_num)));
        lineDataSet.setFillAlpha(110);
        lineDataSet.setCircleRadius(1);
        lineData.addDataSet(lineDataSet);

        binding.lineChart.setData(lineData);
        binding.lineChart.setVisibleXRangeMaximum(entries.get(i).size());
        binding.lineChart.invalidate();
    }

    private void takeScreenshot() throws IOException {
        OutputStream imageOutStream;
        if (Build.VERSION.SDK_INT >= Build.VERSION_CODES.Q) {

            ContentValues values = new ContentValues();
            values.put(MediaStore.Images.Media.DISPLAY_NAME, "image_screenshot.jpg");
            values.put(MediaStore.Images.Media.MIME_TYPE, "image/jpeg");
            values.put(MediaStore.Images.Media.RELATIVE_PATH, Environment.DIRECTORY_PICTURES);

            screenshotUri = getContentResolver().insert(MediaStore.Images.Media.EXTERNAL_CONTENT_URI, values);
            imageOutStream = getContentResolver().openOutputStream(screenshotUri);
        } else {
            String imagesDir = Environment.getExternalStoragePublicDirectory(Environment.DIRECTORY_PICTURES).toString();
            File image = new File(imagesDir, "image_screenshot.jpg");
            imageOutStream = new FileOutputStream(image);

            screenshotUri = Uri.fromFile(image);
        }

        try {
            getViewScreenshot(binding.lineChart).compress(Bitmap.CompressFormat.JPEG, 100, imageOutStream);
        } finally {
            imageOutStream.close();
        }
    }

    @Override
    protected void onSaveInstanceState(@NonNull Bundle outState) {
        super.onSaveInstanceState(outState);
    }

    @Override
    protected void onRestoreInstanceState(@NonNull Bundle savedInstanceState) {
        super.onRestoreInstanceState(savedInstanceState);
    }

    private Bitmap getViewScreenshot(View view) {
        //Define a bitmap with the same size as the view
        Bitmap returnedBitmap = Bitmap.createBitmap(view.getWidth(), view.getHeight(), Bitmap.Config.ARGB_8888);
        //Bind a canvas to it
        Canvas canvas = new Canvas(returnedBitmap);
        //Get the view's background
        Drawable bgDrawable = view.getBackground();
        if (bgDrawable != null) {
            //has background drawable, then draw it on the canvas
            bgDrawable.draw(canvas);
        } else {
            //does not have background drawable, then draw white background on the canvas
            canvas.drawColor(Color.WHITE);
        }
        // draw the view on the canvas
        view.draw(canvas);
        //return the bitmap
        return returnedBitmap;
    }

    private void openScreenshot() {
        Intent intent = new Intent();
        intent.setAction(Intent.ACTION_VIEW);
        intent.setDataAndType(screenshotUri, "image/*");
        startActivity(intent);
    }

    private void renameExcelFile(String filename) throws IOException{
        OutputStream os;
        if (Build.VERSION.SDK_INT >= Build.VERSION_CODES.Q) {

            ContentValues values = new ContentValues();
            values.put(MediaStore.MediaColumns.DISPLAY_NAME, filename + ".xlsx");
            values.put(MediaStore.MediaColumns.RELATIVE_PATH, Environment.DIRECTORY_DOCUMENTS);

            Uri uri = getContentResolver().insert(MediaStore.Files.getContentUri("external"), values);
            os = getContentResolver().openOutputStream(uri);

        } else {
            String filesDir = Environment.getExternalStoragePublicDirectory(Environment.DIRECTORY_DOCUMENTS).toString();
            File doc = new File(filesDir, filename + ".xlsx");
            os = new FileOutputStream(doc);
        }

        try {
            wb.write(os);
            Toast.makeText(MainActivity.this, "Saved!", Toast.LENGTH_SHORT).show();
        } finally {
            os.close();
        }

        /*
        //File documents = Environment.getExternalStoragePublicDirectory(Environment.DIRECTORY_DOCUMENTS);
        File documents = Environment.getExternalStoragePublicDirectory("");

        File folder = new File(documents, "Excel Reader");

        if(!folder.exists()){
            folder.mkdirs();
        }

        File file = new File(folder,   filename + ".xlsx");

        try (FileOutputStream os = new FileOutputStream(file)) {
            wb.write(os);
            Toast.makeText(MainActivity.this, "Saved: " + file.getAbsolutePath(), Toast.LENGTH_SHORT).show();
        } catch (IOException e) {
            Toast.makeText(MainActivity.this, "Error writing " + file + ": " + e.getMessage(), Toast.LENGTH_SHORT).show();
        } catch (Exception e) {
            Toast.makeText(MainActivity.this, "Failed to save file: " + e.getMessage(), Toast.LENGTH_SHORT).show();
        }
         */
    }

    ActivityResultLauncher<Intent> chooseFileActivityResultLauncher = registerForActivityResult(
            new ActivityResultContracts.StartActivityForResult(),
            new ActivityResultCallback<ActivityResult>() {
                @Override
                public void onActivityResult(ActivityResult result) {
                    if (result.getResultCode() == Activity.RESULT_OK) {
                        // There are no request codes
                        Intent data = result.getData();
                        Uri fileUri = null;
                        if (data != null) {
                            fileUri = data.getData();
                        }
                        try{
                            InputStream is = getContentResolver().openInputStream(fileUri);
                            // Create a workbook using the File System
                            wb = new XSSFWorkbook(is);

                            XSSFSheet sheet = wb.getSheetAt(1);

                            // Column 3 .. last Column
                            // Row 0 is header

                            Row header = sheet.getRow(0);

                            labels.clear();
                            entries.clear();

                            for (int i = 3; i < header.getLastCellNum(); i++) {
                                labels.add(header.getCell(i).getStringCellValue());
                                entries.add(new ArrayList<>());
                            }

                            for (int i = 1; i < sheet.getLastRowNum(); i++) {
                                Row value = sheet.getRow(i);
                                for (int j = 3; j < value.getLastCellNum(); j++) {
                                    float cellValue = (float) value.getCell(j).getNumericCellValue();
                                    entries.get(j-3).add(new Entry(i, cellValue));
                                }
                            }

                            currentChartIndex = 0;
                            drawChart(currentChartIndex);

                        }catch (Exception e){
                            Toast.makeText(MainActivity.this,
                                    "Exception: " + e.getMessage(), Toast.LENGTH_SHORT).show();
                            e.printStackTrace();
                        }
                    }
                }
            });
}