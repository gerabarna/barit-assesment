package hu.gerab.barit.assesment;

import static java.util.stream.Collectors.toList;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.TreeMap;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class SummaryBuilder {

    public static final String ASSESSMENT_SHEET_NAME = "ertekeles";
    private static final List<String> expectedHeaders =
            List.of("Név", "jel", "dátum", "barlang", "mozgás", "közösségi viselkedés", "egyéb");

    private final Map<String, List<Assessment>> nameToAssessments = new TreeMap<>();
    private final Map<String, List<Assessment>> signToAssessments = new TreeMap<>();

    public void loadFrom(Path dataDirPath) throws IOException {
        List<Path> files = Files.find(dataDirPath, 1,
                        (path, attributes) -> attributes.isRegularFile() && path.getFileName().toString().endsWith(".xlsx"))
                .collect(toList());

        for (Path filePath : files) {
            System.out.println("Processing path=" + filePath);
            FileInputStream file = new FileInputStream(filePath.toFile());
            Workbook workbook = new XSSFWorkbook(file);
            Sheet sheet = workbook.getSheet(ASSESSMENT_SHEET_NAME);

            Iterator<Row> iterator = sheet.iterator();
            iterator.next(); // skip over header

            while (iterator.hasNext()) {
                Row row = iterator.next();
                Assessment assessment = new Assessment(row);

                String name = assessment.getName();
                if (name == null || name.isBlank()) {
                    String sign = assessment.getSign();
                    if (sign != null && !sign.isBlank()) {
                        signToAssessments.computeIfAbsent(sign, s -> new ArrayList<>()).add(assessment);
                    }
                } else {
                    nameToAssessments.computeIfAbsent(name, n -> new ArrayList<>()).add(assessment);
                }
            }
            //for (List<Assessment> assessments : nameToAssessments.values()) {
            //    for (Assessment assessment : assessments) {
            //        System.out.println(assessment);
            //    }
            //}
            //for (List<Assessment> assessments : signToAssessments.values()) {
            //    for (Assessment assessment : assessments) {
            //        System.out.println(assessment);
            //    }
            //}
        }
    }

    public void writeTo(Path outputPath) throws IOException {
        try (Workbook workbook = new XSSFWorkbook();
             FileOutputStream outputStream = new FileOutputStream(outputPath.toFile())) {

            Sheet sheet = workbook.createSheet(ASSESSMENT_SHEET_NAME);
            Row header = sheet.createRow(0);

            CreationHelper createHelper = workbook.getCreationHelper();
            short format = createHelper.createDataFormat().getFormat("yyyy.mm.dd");
            CellStyle dateStyle = workbook.createCellStyle();
            dateStyle.setDataFormat(format);
            sheet.setDefaultColumnStyle(2, dateStyle);

            CellStyle headerStyle = workbook.createCellStyle();
            headerStyle.setFillBackgroundColor(IndexedColors.LIGHT_BLUE.getIndex());
            header.setRowStyle(headerStyle);

            for (int i = 0; i < expectedHeaders.size(); i++) {
                String expectedHeader = expectedHeaders.get(i);
                Cell headerCell = header.createCell(i);
                headerCell.setCellValue(expectedHeader);
            }

            int rowCounter = 1; // 0 is the header
            for (List<Assessment> assessments : nameToAssessments.values()) {
                for (Assessment assessment : assessments) {
                    assessment.fillRow(sheet.createRow(rowCounter++));
                }
            }
            for (List<Assessment> assessments : signToAssessments.values()) {
                for (Assessment assessment : assessments) {
                    assessment.fillRow(sheet.createRow(rowCounter++));
                }
            }
            System.out.println(rowCounter);
            workbook.write(outputStream);
        }
    }

}
