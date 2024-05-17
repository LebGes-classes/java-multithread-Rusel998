package org.example;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

public class Task {
    private String name;
    private int idWorker;
    private int hours;
    private boolean status;

    public static final String TASKS_FILE = "tasks.xlsx";

    public Task(String name, int hours) {
        this.name = name;
        this.hours = hours;
        this.status = true;
    }

    public int getIdWorker() {
        return idWorker;
    }

    public void setIdWorker(int idWorker) {
        this.idWorker = idWorker;
    }

    public int getHours() {
        return hours;
    }

    public String getName() {
        return name;
    }

    public boolean isCompleted() {
        return !status;
    }

    public void complete() {
        this.status = false;
    }

    public static void addTaskToTable(Task task) {
        try (FileInputStream fis = new FileInputStream(TASKS_FILE);
             Workbook workbook = new XSSFWorkbook(fis);
             FileOutputStream fos = new FileOutputStream(TASKS_FILE)) {

            Sheet sheet = workbook.getSheetAt(0);
            int lastRowNum = sheet.getLastRowNum();
            Row newRow = sheet.createRow(lastRowNum + 1);

            newRow.createCell(0).setCellValue(task.idWorker);
            newRow.createCell(1).setCellValue(task.hours);
            newRow.createCell(2).setCellValue(task.status);

            workbook.write(fos);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}

