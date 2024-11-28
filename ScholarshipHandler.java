import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.*;

class Student {
    private int id;
    private String name;
    private String group;
    private double currentScholarship;
    private double gpa;
    private String faculty;
    private double newScholarship;

    public Student(int id, String name, String group, double currentScholarship, double gpa, String faculty) {
        this.id = id;
        this.name = name;
        this.group = group;
        this.currentScholarship = currentScholarship;
        this.gpa = gpa;
        this.faculty = faculty;
    }

    public int getId() {
        return id;
    }

    public void setId(int id) {
        this.id = id;
    }

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public String getGroup() {
        return group;
    }

    public void setGroup(String group) {
        this.group = group;
    }

    public double getCurrentScholarship() {
        return currentScholarship;
    }

    public void setCurrentScholarship(double currentScholarship) {
        this.currentScholarship = currentScholarship;
    }

    public double getGpa() {
        return gpa;
    }

    public void setGpa(double gpa) {
        this.gpa = gpa;
    }

    public String getFaculty() {
        return faculty;
    }

    public void setFaculty(String faculty) {
        this.faculty = faculty;
    }

    public double getNewScholarship() {
        return newScholarship;
    }

    public void setNewScholarship(double newScholarship) {
        this.newScholarship = newScholarship;
    }
}

public class ScholarshipHandler {

    public static List<Student> readData(String filePath) throws IOException {
        List<Student> students = new ArrayList<>();
        FileInputStream file = new FileInputStream(new File(filePath));
        Workbook workbook = new XSSFWorkbook(file);
        Sheet sheet = workbook.getSheetAt(0);

        for (Row row : sheet) {
            if (row.getRowNum() == 0) continue;

            try {
                int id = (int) row.getCell(0).getNumericCellValue();
                String name = row.getCell(1).getStringCellValue();
                String group = row.getCell(2).getStringCellValue();
                double currentScholarship = row.getCell(3).getNumericCellValue();
                double gpa = row.getCell(4).getNumericCellValue();
                String faculty = row.getCell(5).getStringCellValue();

                Student student = new Student(id, name, group, currentScholarship, gpa, faculty);
                students.add(student);
            } catch (Exception e) {
                System.out.println("Ошибка чтения данных для строки " + (row.getRowNum() + 1));
            }
        }
        workbook.close();
        return students;
    }

    public static void calculateNewScholarships(List<Student> students) {
        for (Student student : students) {
            double newScholarship = student.getCurrentScholarship();

            if ("Engineering".equals(student.getFaculty()) && student.getGpa() > 2.4) {
                newScholarship = newScholarship * 1.10;
            } else if ("Economics".equals(student.getFaculty()) && student.getGpa() > 2.4) {
                newScholarship = newScholarship * 1.15;
            } else if ("Philosophy".equals(student.getFaculty()) && student.getGpa() > 2.2) {
                newScholarship = newScholarship * 1.05;
            } else if ("Marketing".equals(student.getFaculty()) && student.getGpa() > 2.5) {
                newScholarship = newScholarship * 1.08;
            }

            student.setNewScholarship(newScholarship);
        }
    }

    public static void writeUpdatedData(List<Student> students, String filePath) throws IOException {
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("Updated Students");

        Row headerRow = sheet.createRow(0);
        headerRow.createCell(0).setCellValue("ID");
        headerRow.createCell(1).setCellValue("Name");
        headerRow.createCell(2).setCellValue("Group");
        headerRow.createCell(3).setCellValue("Current Scholarship");
        headerRow.createCell(4).setCellValue("GPA");
        headerRow.createCell(5).setCellValue("Faculty");
        headerRow.createCell(6).setCellValue("New Scholarship");

        int rowIndex = 1;
        for (Student student : students) {
            Row row = sheet.createRow(rowIndex++);
            row.createCell(0).setCellValue(student.getId());
            row.createCell(1).setCellValue(student.getName());
            row.createCell(2).setCellValue(student.getGroup());
            row.createCell(3).setCellValue(student.getCurrentScholarship());
            row.createCell(4).setCellValue(student.getGpa());
            row.createCell(5).setCellValue(student.getFaculty());
            row.createCell(6).setCellValue(student.getNewScholarship());
        }

        FileOutputStream fileOut = new FileOutputStream(filePath);
        workbook.write(fileOut);
        fileOut.close();
        workbook.close();
    }

    public static void main(String[] args) throws IOException {
        List<Student> students = readData("C:\\Users\\Сержанбек\\IdeaProjects\\untitled7\\src\\students.xlsx");
        calculateNewScholarships(students);
        writeUpdatedData(students, "C:\\Users\\Сержанбек\\IdeaProjects\\untitled7\\src\\updated_students.xlsx");
    }
}
