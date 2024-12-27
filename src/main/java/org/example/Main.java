package org.example;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.time.LocalDate;
import java.time.format.TextStyle;
import java.util.*;

public class Main {

    static class TeamMember {
        String name;
        String team;
        boolean prefersMorning;
        Set<Integer> offDays = new HashSet<>();
        int shiftCount = 0;

        public TeamMember(String name, String team, boolean prefersMorning) {
            this.name = name;
            this.team = team;
            this.prefersMorning = prefersMorning;
        }
    }

    static class Shift {
        String date;
        String dayOfWeek;
        String time; // "7am-1pm" or "1pm-7pm"
        String nwTeamMember;
        String csTeamMember;

        public Shift(String date, String dayOfWeek, String time, String nwTeamMember, String csTeamMember) {
            this.date = date;
            this.dayOfWeek = dayOfWeek;
            this.time = time;
            this.nwTeamMember = nwTeamMember;
            this.csTeamMember = csTeamMember;
        }

        @Override
        public String toString() {
            return String.format("%s (%s) (%s): NW: %s, CS: %s", date, dayOfWeek, time, nwTeamMember, csTeamMember);
        }

        public String toCSV() {
            return String.format("%s,%s,%s,%s,%s", date, dayOfWeek, time, nwTeamMember, csTeamMember);
        }
    }

    public static void main(String[] args) throws IOException {
        // Create the NW and CS teams
        List<TeamMember> nwTeam = Arrays.asList(
                new TeamMember("Sajad", "NW", false),
                new TeamMember("Harish", "NW", false),
                new TeamMember("Paul", "NW", true),
                new TeamMember("Mahesh", "NW", false),
                new TeamMember("Mithra", "NW", false),
                new TeamMember("Matthew", "NW", false)
        );

        List<TeamMember> csTeam = Arrays.asList(
                new TeamMember("Palan", "CS", false),
                new TeamMember("Victor", "CS", false),
                new TeamMember("Divya", "CS", true),
                new TeamMember("Gaston", "CS", true),
                new TeamMember("Suraj", "CS", false),
                new TeamMember("Dhiren", "CS", false),
                new TeamMember("Sacha", "CS", false)
        );

        // Example off-days setup
        nwTeam.get(1).offDays.addAll(Arrays.asList(15)); // Harish off on 15th
        nwTeam.get(2).offDays.addAll(Arrays.asList(3)); // Paul off on 3rd
        csTeam.get(3).offDays.addAll(Arrays.asList(21, 22, 23, 24, 27, 28, 29, 30, 31)); // Gaston off 21st to 31st

        // Generate the monitoring rota for January
        List<Shift> shifts = generateRota(nwTeam, csTeam);

        // Write the shifts to an Excel file
        writeShiftsToExcel(shifts);

        // Output the shift counts for each person
        System.out.println("\nShift counts for each person:");
        printShiftCounts(nwTeam, csTeam);
    }

    private static List<Shift> generateRota(List<TeamMember> nwTeam, List<TeamMember> csTeam) {
        List<Shift> rota = new ArrayList<>();
        Set<String> scheduledNW = new HashSet<>();
        Set<String> scheduledCS = new HashSet<>();

        // Generate the rota for January 2025 (31 days)
        for (int day = 1; day <= 31; day++) {
            // Skip the bank holiday on 1st January
            if (day == 1) {
                continue;
            }

            // Determine the date and the weekday of the current day
            String date = "January " + day;
            String dayOfWeek = getDayOfWeek(day);  // Get the day of the week (e.g., Monday, Tuesday, etc.)

            // Skip weekends (Saturday and Sunday)
            if (dayOfWeek.equals("Saturday") || dayOfWeek.equals("Sunday")) {
                continue;
            }

            // Assign shifts for both morning and afternoon for the current day
            for (String time : Arrays.asList("7am-1pm", "1pm-7pm")) {
                TeamMember nw = findAvailableTeamMember(nwTeam, day, time.equals("7am-1pm"), scheduledNW);
                TeamMember cs = findAvailableTeamMember(csTeam, day, time.equals("7am-1pm"), scheduledCS);

//                Finds an available NW and CS team member for morning (7am-1pm) and afternoon (1pm-7pm) shifts.
//                Skips members who:
//                 - Prefer different times.
//                 - Have reached the shift limit.
//                 - Are off on that day.

                if (nw != null && cs != null) {
                    rota.add(new Shift(date, dayOfWeek, time, nw.name, cs.name));
                    scheduledNW.add(nw.name);
                    scheduledCS.add(cs.name);
                    nw.shiftCount++; // Update shift count for NW team member
                    cs.shiftCount++; // Update shift count for CS team member
                }
            }

            // Reset daily constraints after assigning shifts
            scheduledNW.clear();
            scheduledCS.clear();
        }

        return rota;
    }


    private static String getDayOfWeek(int day) {
        // Using LocalDate API to get the day of the week (0 = Sunday, 6 = Saturday)
        LocalDate date = LocalDate.of(2025, 1, day); // January
        return date.getDayOfWeek().getDisplayName(TextStyle.FULL, Locale.ENGLISH);
    }

    private static TeamMember findAvailableTeamMember(List<TeamMember> team, int day, boolean isMorning, Set<String> scheduled) {
        // Find a team member who is available to work the shift
        List<TeamMember> availableMembers = new ArrayList<>();

        // Iterate over team members and add available ones to the list
        for (TeamMember member : team) {
            // Skip if already scheduled today or has max shifts (7 shifts per month)
            if (scheduled.contains(member.name) || member.shiftCount >= 8) continue;

            // Skip if they are off today or the day before
            if (member.offDays.contains(day) || member.offDays.contains(day - 1)) continue;

            // Skip if they prefer afternoon but it's a morning shift
            if (!isMorning && member.prefersMorning) continue;

            availableMembers.add(member);
        }

        // Log available members
        if (availableMembers.isEmpty()) {
            System.out.println("No available members for day " + day + " shift.");
        }

        // If there are available members, pick one randomly to balance the shift count
        if (!availableMembers.isEmpty()) {
            // Sort available members by their current shift count to balance the assignment
            availableMembers.sort(Comparator.comparingInt(a -> a.shiftCount));

            return availableMembers.get(0); // Assign the one with the least shifts
        }

        return null;
    }

    private static void writeShiftsToExcel(List<Shift> shifts) throws IOException {
        // Create a workbook and a sheet
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("Shift Rota");

        // Create a header row
        Row headerRow = sheet.createRow(0);
        headerRow.createCell(0).setCellValue("Date");
        headerRow.createCell(1).setCellValue("Day of Week");
        headerRow.createCell(2).setCellValue("Shift Time");
        headerRow.createCell(3).setCellValue("NW Team Member");
        headerRow.createCell(4).setCellValue("CS Team Member");

        // Add styles for formatting
        CellStyle headerStyle = workbook.createCellStyle();
        headerStyle.setFillForegroundColor(IndexedColors.LIGHT_BLUE.getIndex());
        headerStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        headerStyle.setAlignment(HorizontalAlignment.CENTER);

        // Apply header style
        for (Cell cell : headerRow) {
            cell.setCellStyle(headerStyle);
        }

        // Populate the sheet with shift data
        int rowIndex = 1;
        for (Shift shift : shifts) {
            Row row = sheet.createRow(rowIndex++);
            row.createCell(0).setCellValue(shift.date);
            row.createCell(1).setCellValue(shift.dayOfWeek);
            row.createCell(2).setCellValue(shift.time);
            row.createCell(3).setCellValue(shift.nwTeamMember);
            row.createCell(4).setCellValue(shift.csTeamMember);
        }

        // Write the workbook to a file
        try (FileOutputStream fileOut = new FileOutputStream("shift_rota_january_2025.xlsx")) {
            workbook.write(fileOut);
        }
        workbook.close();

        System.out.println("\nShifts have been written to shift_rota_january_2025.xlsx");
    }

    private static void printShiftCounts(List<TeamMember> nwTeam, List<TeamMember> csTeam) {
        // Print shift counts for each team member
        System.out.println("\nShift counts for each person:");
        for (TeamMember member : nwTeam) {
            System.out.println(member.name + " (" + member.team + "): " + member.shiftCount + " shifts");
        }
        for (TeamMember member : csTeam) {
            System.out.println(member.name + " (" + member.team + "): " + member.shiftCount + " shifts");
        }
    }
}
