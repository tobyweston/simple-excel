package bad.robot.excel;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;

import static java.io.File.createTempFile;
import static org.apache.poi.util.IOUtils.copy;

public class OutputExcelWorkbook {

    public static void writeWorkbookToTemporaryFile(ExcelWorkbook workbook, String filename) throws IOException {
        File file = createTempFile(filename, ".xls");
        System.out.println("outputted xls: " + file.getAbsolutePath());
        copy(workbook.getInputStream(), new FileOutputStream(file));
    }

}
