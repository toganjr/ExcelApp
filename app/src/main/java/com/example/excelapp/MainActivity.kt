package com.example.excelapp

import android.Manifest
import android.annotation.SuppressLint
import android.app.Activity
import android.content.ContentValues.TAG
import android.content.Context
import android.content.Intent
import android.net.Uri
import android.os.Build
import androidx.appcompat.app.AppCompatActivity
import android.os.Bundle
import android.provider.DocumentsContract
import android.util.Log
import android.widget.TextView
import androidx.annotation.RequiresApi
import org.apache.poi.hssf.usermodel.HSSFWorkbook
import java.io.File
import java.io.FileInputStream
import java.io.IOException
import androidx.core.app.ActivityCompat.startActivityForResult

import androidx.core.app.ActivityCompat

import android.content.pm.PackageManager
import android.provider.MediaStore

import androidx.core.content.ContextCompat
import android.widget.Toast
import androidx.core.content.PackageManagerCompat

import androidx.core.content.PackageManagerCompat.LOG_TAG
import org.apache.poi.ss.usermodel.*
import org.apache.poi.ss.usermodel.CellType.*
import org.apache.poi.util.IOUtils
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import java.io.FileOutputStream


class MainActivity : AppCompatActivity() {

    private val MY_REQUEST_CODE_PERMISSION = 1000
    private val MY_RESULT_CODE_FILECHOOSER = 2000
    private val activity = this
    val arrayname2 = arrayOf<String>("","","","","")

    override fun onCreate(savedInstanceState: Bundle?) {
        super.onCreate(savedInstanceState)
        setContentView(R.layout.activity_main)

        // get reference to button
        val btn_click_me = findViewById(R.id.button) as TextView
        val text = findViewById(R.id.text) as TextView
    // set on-click listener
        btn_click_me.setOnClickListener {
            askPermissionAndBrowseFile()
            for (element in arrayname2) {
                text.text = "${text.text}+${element}"
            }
        }

    }

    private fun askPermissionAndBrowseFile() {
        // With Android Level >= 23, you have to ask the user
        // for permission to access External Storage.
        if (Build.VERSION.SDK_INT >= Build.VERSION_CODES.Q) {
            ActivityCompat.requestPermissions(
                this, arrayOf(
                    Manifest.permission.READ_EXTERNAL_STORAGE
                ), 2
            )
        } else {
            ActivityCompat.requestPermissions(
                this, arrayOf(
                    Manifest.permission.WRITE_EXTERNAL_STORAGE,
                    Manifest.permission.READ_EXTERNAL_STORAGE
                ), 2
            )
        }
        this.doBrowseFile()
    }

    private fun doBrowseFile() {
        try {
            if (Build.VERSION.SDK_INT >= Build.VERSION_CODES.Q) {
                intent = Intent(Intent.ACTION_PICK,MediaStore.Images.Media.EXTERNAL_CONTENT_URI);
                intent.setType("*/*");
                intent.putExtra(Intent.EXTRA_LOCAL_ONLY, true);
                intent.setAction(Intent.ACTION_GET_CONTENT);
                activity.startActivityForResult(Intent.createChooser(intent, "Select File"), MY_RESULT_CODE_FILECHOOSER);
            }else {
                intent = Intent();
                intent.setType("*/*");
                intent.putExtra(Intent.EXTRA_LOCAL_ONLY, true);
                intent.setAction(Intent.ACTION_GET_CONTENT);
                activity.startActivityForResult(Intent.createChooser(intent, "Select File"), MY_RESULT_CODE_FILECHOOSER);
            }
        } catch (e : Exception) {
            e.printStackTrace();
        }

//        ar chooseFileIntent = Intent(Intent.ACTION_GET_CONTENT)
//        chooseFileIntent.type = "*/*"
//        // Only return URIs that can be opened with ContentResolver
//        chooseFileIntent.addCategory(Intent.CATEGORY_OPENABLE)
//        chooseFileIntent = Intent.createChooser(chooseFileIntent, "Choose a file")
//        startActivityForResult(chooseFileIntent, MY_RESULT_CODE_FILECHOOSER)
    }


    @SuppressLint("RestrictedApi")
    override fun onRequestPermissionsResult(
        requestCode: Int,
        permissions: Array<out String>,
        grantResults: IntArray
    ) {
        super.onRequestPermissionsResult(requestCode, permissions, grantResults)
        when (requestCode) {
            MY_REQUEST_CODE_PERMISSION -> {

                // Note: If request is cancelled, the result arrays are empty.
                // Permissions granted (CALL_PHONE).
                if (grantResults.isNotEmpty()
                    && grantResults[0] == PackageManager.PERMISSION_GRANTED) {

                    Log.i( LOG_TAG,"Permission granted!");
                    Toast.makeText(baseContext, "Permission granted!", Toast.LENGTH_SHORT).show();

                    this.doBrowseFile();
                }
                // Cancelled or denied.
                else {
                    Log.i(LOG_TAG,"Permission denied!");
                    Toast.makeText(baseContext, "Permission denied!", Toast.LENGTH_SHORT).show();
                }
            }
        }
    }

    @SuppressLint("RestrictedApi")
    override fun onActivityResult(requestCode: Int, resultCode: Int, data: Intent?) {
        when (requestCode) {
            MY_RESULT_CODE_FILECHOOSER -> {
                if (resultCode == Activity.RESULT_OK ) {
                    if(data != null)  {
                        val fileUri = data.data
                        Log.i(LOG_TAG, "Uri: " + fileUri);

                        var filePath: String? = null;
                        try {
                            filePath = FileUtils.getPath(baseContext,fileUri!!);
                        } catch (e: Exception) {
                            Log.e(LOG_TAG,"Error: " + e);
                            Toast.makeText(baseContext, "Error: " + e, Toast.LENGTH_SHORT).show();
                        }
                        val parcelFileDescriptor = baseContext.contentResolver.openFileDescriptor(fileUri!!, "r", null)
//                        val file = File(filePath!!)
//                        readExcelFromStorage(baseContext,file.toString())
                        parcelFileDescriptor?.let {
                            val inputStream = FileInputStream(parcelFileDescriptor.fileDescriptor)
                            val file = File(filePath!!)
                            val copyFile = File(baseContext.cacheDir, file.name)
                            val outputStream = FileOutputStream(copyFile)
                            IOUtils.copy(inputStream,outputStream)
                            readExcelFromStorage(baseContext,copyFile.toString())
                        }

                    }
                }
            }
        }
        super.onActivityResult(requestCode, resultCode, data)
    }

    fun readExcelFromStorage(context: Context, fileName: String) {
        val file = File(fileName);
        var fileInputStream : FileInputStream? = null

        try {
            fileInputStream = FileInputStream(file)
            Log.e(TAG, "Reading from Excel$file")

//            // Create instance having reference to .xls file
//            val workbook = HSSFWorkbook(fileInputStream)

            // Create instance having reference to .xlsx file
//            val workbook = XSSFWorkbook(fileInputStream)

            //Create instance having reference to .xlsx and .xls file
            val workbook = WorkbookFactory.create(fileInputStream)

            // Fetch sheet at position 'i' from the workbook
            val sheet = workbook.getSheetAt(0)

            var po = 0
            // Iterate through each row
            for (row: Row in sheet) {

                if (row.rowNum > 0) run {
                    val cellIterator: Iterator<Cell> = row.cellIterator()


                    while(cellIterator.hasNext()) {
                        var cell: Cell = cellIterator.next()

                        when (cell.cellType) {
                            Cell.CELL_TYPE_NUMERIC ->  {
                                arrayname2[po] = cell.numericCellValue.toString()
                                po++
                            }
                            Cell.CELL_TYPE_STRING -> {
                                arrayname2[po] = cell.stringCellValue
                                po++
                            }
                        }
                        print("ISI "+arrayname2)
                    }
                }
            }
        } catch (e: IOException) {
            Log.e(TAG, "Error Reading Exception: ", e);
        } catch (e: Exception) {
            Log.e(TAG, "Failed to read file due to Exception: ", e);
        } finally {
            try {
                fileInputStream?.close()
            } catch (ex: Exception) {
                ex.printStackTrace();
            }
        }
    }
}