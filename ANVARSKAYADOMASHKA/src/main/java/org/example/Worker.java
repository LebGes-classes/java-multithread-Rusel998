package org.example;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.LinkedList;
import java.util.List;

public class Worker implements Runnable {
    public static LinkedList<Worker> workerList = new LinkedList<>();

    private String name;
    private final int id;
    private int workedHours;
    private int afkHours;
    private LinkedList<Task> tasks;
    private final int MAX_WORK_HOURS = 8;
    private int efficiency;

    private static final String WORKERS_FILE = "workers.xlsx";
    private static final String TASKS_FILE = "tasks.xlsx";

    public Worker(String name, int id, int workedHours, int afkHours, int efficiency) {
        this.name = name;
        this.id = id;
        this.tasks = new LinkedList<>();
        this.workedHours = workedHours;
        this.afkHours = afkHours;
        this.efficiency = efficiency;
    }

    public String getName() {
        return name;
    }

    public int getId() {
        return id;
    }

    public int getWorkedHours() {
        return workedHours;
    }

    public int getAfkHours() {
        return afkHours;
    }

    public int getEfficiency() {
        return efficiency;
    }

    public List<Task> getTasks() {
        return tasks;
    }

    public void addTask(Task task) {
        this.tasks.add(task);
    }

    @Override
    public void run() {
        System.out.println("Worker " + name + " started working.");

        for (Task task : tasks) {
            System.out.println("Worker " + name + " started task " + task.getName());

            int remainingHours = task.getHours();
            while (remainingHours > 0) {
                int hoursToWork = Math.min(remainingHours, MAX_WORK_HOURS);

                for (int i = 0; i < hoursToWork; i++) {
                    workedHours++;
                    remainingHours--;

                    System.out.println("Worker " + name + " has worked on task " + task.getName() + " for 1 hour. Remaining hours: " + remainingHours);

                    try {
                        Thread.sleep(1000);
                    } catch (InterruptedException e) {
                        throw new RuntimeException(e);
                    }
                }

                if (remainingHours > 0) {
                    System.out.println("Task " + task.getName() + " will continue tomorrow. Remaining hours: " + remainingHours);
                    afkHours += 16;
                    try {
                        Thread.sleep(2000);
                    } catch (InterruptedException e) {
                        throw new RuntimeException(e);
                    }
                }
            }

            task.complete();
            System.out.println("Task " + task.getName() + " completed by Worker " + name);
        }

        efficiency = (workedHours / afkHours) * 100;
        System.out.println("Worker " + name + " has completed all tasks.");
    }

    public static List<Worker> readWorkersFromExcel() {
        List<Worker> workers = new ArrayList<>();

        try (FileInputStream fis = new FileInputStream(WORKERS_FILE); Workbook workbook = new XSSFWorkbook(fis)) {
            Sheet sheet = workbook.getSheetAt(0);

            for (Row row : sheet) {
                String name = row.getCell(0).getStringCellValue();
                int id = (int) row.getCell(1).getNumericCellValue();
                int workedHours = (int) row.getCell(2).getNumericCellValue();
                int afkHours = (int) row.getCell(3).getNumericCellValue();
                int efficiency = (int) row.getCell(4).getNumericCellValue();
                workers.add(new Worker(name, id, workedHours, afkHours, efficiency));
            }
        } catch (IOException e) {
            e.printStackTrace();
        }

        return workers;
    }

    public static void writeWorkersToExcel(List<Worker> workers) {
        try (Workbook workbook = new XSSFWorkbook(); FileOutputStream fos = new FileOutputStream(WORKERS_FILE)) {
            Sheet sheet = workbook.createSheet("Workers");

            int rowNum = 0;
            for (Worker worker : workers) {
                Row row = sheet.createRow(rowNum++);
                row.createCell(0).setCellValue(worker.getName());
                row.createCell(1).setCellValue(worker.getId());
                row.createCell(2).setCellValue(worker.getWorkedHours());
                row.createCell(3).setCellValue(worker.getAfkHours());
                row.createCell(4).setCellValue(worker.getEfficiency());
            }

            workbook.write(fos);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    public static void assignTasksToWorkers(List<Worker> workers) {
        try (FileInputStream fis = new FileInputStream(TASKS_FILE); Workbook workbook = new XSSFWorkbook(fis)) {
            Sheet sheet = workbook.getSheetAt(0);

            for (Row row : sheet) {
                int workerId = (int) row.getCell(0).getNumericCellValue();
                String taskDescription = row.getCell(1).getStringCellValue();
                int taskHours = (int) row.getCell(2).getNumericCellValue();

                Worker worker = findWorkerById(workers, workerId);
                if (worker != null) {
                    Task task = new Task(taskDescription, taskHours);
                    task.setIdWorker(workerId);
                    worker.addTask(task);
                } else {
                    System.out.println("Worker with id " + workerId + " not found.");
                }
            }
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    public static void writeTasksToExcel(List<Worker> workers) {
        try (Workbook workbook = new XSSFWorkbook(); FileOutputStream fos = new FileOutputStream(TASKS_FILE)) {
            Sheet sheet = workbook.createSheet("Tasks");

            int rowNum = 0;
            for (Worker worker : workers) {
                for (Task task : worker.getTasks()) {
                    Row row = sheet.createRow(rowNum++);
                    row.createCell(0).setCellValue(worker.getId());
                    row.createCell(1).setCellValue(task.getName());
                    row.createCell(2).setCellValue(task.getHours());
                }
            }

            workbook.write(fos);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private static Worker findWorkerById(List<Worker> workers, int id) {
        for (Worker worker : workers) {
            if (worker.getId() == id) {
                return worker;
            }
        }
        return null;
    }
}