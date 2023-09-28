package com.example;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.sql.Time;
import java.time.Duration;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.LocalTime;
import java.time.format.DateTimeFormatter;
import java.util.HashMap;
import java.util.Map;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Employee {
    public static void main(String[] args) {
        // Create a map to store the number of days each employee worked
            Map<String, Integer> dayCount = new HashMap<>();
        try {
            // Open the Excel file
            FileInputStream file = new FileInputStream(new File("java\\demo\\src\\main\\java\\com\\example\\Assignment_Timecard.xlsx"));

            Workbook workbook = new XSSFWorkbook(file);

            Sheet sheet = workbook.getSheetAt(0);

            // map to store the last day each employee worked
            Map<String, LocalDateTime> lastDay = new HashMap<>();

            //  map to store the last day each employee worked
            Map<String, LocalDate> countDayTrack = new HashMap<>();
            

            for (Row row : sheet) {
                // Skip the header rows
                if (row.getRowNum() < 2) {
                    continue;
                }

                // Get the employee name and position
                String name = row.getCell(7).getStringCellValue();
                String position = row.getCell(0).getStringCellValue();

                // Get the date and time of the current row
                row.getCell(2).setCellType(CellType.NUMERIC);
                LocalDateTime dateTime = row.getCell(2).getLocalDateTimeCellValue();
                
                // count the number of days worked
                workDaysCount(name, dateTime.toLocalDate(), countDayTrack,dayCount);

                // Check if the employee has less than 10 hours of time between shifts but greater than 1 hour
                if (hoursBetweenShifts(name, dateTime, lastDay)) {
                    System.out.println(name +" - "+ position+ " : has less than 10 hours and more than 1 hour between shifts ");
                }

                // Check if the employee has worked for more than 14 hours in a single shift
                if (workedMoreThan14(name,row)) {
                    System.out.println(name +" - "+ position + " : has worked for more than 14 hours in a single shift" );
                }

                // Update the last day each employee worked
                row.getCell(3).setCellType(CellType.NUMERIC);
                LocalDateTime dayEnd = row.getCell(3).getLocalDateTimeCellValue();
                lastDay.put(name, dayEnd);
            }

            // Close the workbook and file
            workbook.close();
            file.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
        finally{
            // print the employees who worked for 7 days in a row
            for (Map.Entry<String, Integer> entry : dayCount.entrySet()) {
                if(entry.getValue() >= 7){
                    System.out.println(entry.getKey() + " has worked for 7 days in a row");
                }
            }
        }
    }

   // method to count the number of days worked
    private static void workDaysCount(String name, LocalDate dateTime, Map<String, LocalDate> tracker, Map<String, Integer> dayCount) {
        
        if (!tracker.containsKey(name)) {
            tracker.put(name, dateTime);
            return;
        }
        LocalDate lastDate = tracker.get(name);
        // if the last date is not null and the current date is the next day of the last date
        if (lastDate != null && lastDate.plusDays(1).isEqual(dateTime)) {
            dayCount.put(name, dayCount.getOrDefault(name, 0) + 1);
            tracker.put(name, dateTime);
        }
        
    }



    private static boolean hoursBetweenShifts(String name, LocalDateTime dateTime, Map<String, LocalDateTime> lastDay) {
        LocalDateTime exiTime = lastDay.get(name);
        
        // if the last day  is not null and the current date is the last day are same  as employee have shifts in one day
        if (exiTime != null && dateTime.toLocalDate().isEqual(exiTime.toLocalDate())) {
            Duration duration = Duration.between(exiTime, dateTime);
            // if the duration is less than 10 hours and greater than 1 hour
            if (duration.toHours() > 1 && duration.toHours() < 10) {
                return true;
            }
        }
        return false;
    }

    private static boolean workedMoreThan14(String name,  Row row) {
        // get the timcar hours cell
        Cell duration = row.getCell(4);

        // to handle numeric cells and convert them into string
        duration.setCellType(CellType.STRING);

        // to handle empty cells
        if (duration.getStringCellValue().equals("")){
            return false;
        }

        DateTimeFormatter formatter = DateTimeFormatter.ofPattern("H:mm");
        LocalTime time = LocalTime.parse(duration.getStringCellValue(), formatter);
        if (time.getHour() > 14) {
            return true;
        }else{
            return false;
        }       
    }
}