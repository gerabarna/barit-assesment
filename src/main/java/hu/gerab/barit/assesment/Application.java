package hu.gerab.barit.assesment;


import java.io.IOException;
import java.nio.file.Path;

public class Application {

    public static void main(String[] args) throws IOException {
        Path dataDirPath = Path.of("/home/gerab/temp/barit/ertekeles/2023_09_28");
        Path outputPath = Path.of("/home/gerab/temp/barit/ertekeles/2023_09_28_summary.xlsx");

        final SummaryBuilder summaryBuilder = new SummaryBuilder();
        summaryBuilder.loadFrom(dataDirPath);
        summaryBuilder.writeTo(outputPath);
    }

}