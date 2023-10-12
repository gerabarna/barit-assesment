package hu.gerab.barit.assesment;

import java.time.LocalDate;
import java.time.ZoneId;
import java.time.format.DateTimeFormatter;
import lombok.Data;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;

@Data
class Assessment {
    public static final String DATE_PATERN = "yyyy.MM.dd";
    public static final DateTimeFormatter dateFormatter = DateTimeFormatter.ofPattern(DATE_PATERN);

    private String name;
    private String sign;
    private LocalDate date;
    private String caveName;
    private String movement;
    private String sociability;
    private String etc;

    public Assessment(Row row) {
        Cell nameCell = row.getCell(0);
        Cell signCell = row.getCell(1);
        Cell dateCell = row.getCell(2);
        Cell caveNameCell = row.getCell(3);
        Cell movementCell = row.getCell(4);
        Cell sociabilityCell = row.getCell(5);
        Cell etcCell = row.getCell(6);

        name = nameCell == null ? null : nameCell.getStringCellValue();
        sign = signCell == null ? null : signCell.getStringCellValue();
        date = parseLocalDate(dateCell);
        if (date != null && date.getYear() < 2023) {
            date = date.plusYears(2023 - date.getYear());
        }
        caveName = caveNameCell == null ? null : caveNameCell.getStringCellValue();
        movement = movementCell == null ? null : movementCell.getStringCellValue();
        sociability = sociabilityCell == null ? null : sociabilityCell.getStringCellValue();
        etc = etcCell == null ? null : etcCell.getStringCellValue();
    }

    public Row fillRow(Row row) {
        int columnCounter = 0;
        row.createCell(columnCounter++).setCellValue(getName());
        row.createCell(columnCounter++).setCellValue(getSign());
        row.createCell(columnCounter++).setCellValue(getDate());
        row.createCell(columnCounter++).setCellValue(getCaveName());
        row.createCell(columnCounter++).setCellValue(getMovement());
        row.createCell(columnCounter++).setCellValue(getSociability());
        row.createCell(columnCounter++).setCellValue(getEtc());
        return row;
    }

    private LocalDate parseLocalDate(Cell dateCell) {
        if (dateCell == null) {
            return null;
        } else {
            return switch (dateCell.getCellType()) {
                case STRING -> {
                    final String stringCellValue = dateCell.getStringCellValue();
                    if (stringCellValue.isBlank()) {
                        yield null;
                    }
                    yield LocalDate.parse(stringCellValue, dateFormatter);
                }
                case NUMERIC -> LocalDate.ofInstant(dateCell.getDateCellValue().toInstant(), ZoneId.of("CET"));
                default ->
                        throw new IllegalArgumentException("Cannot parse date from cellType=" + dateCell.getCellType());
            };
        }
    }
}
