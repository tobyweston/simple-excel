package bad.robot.excel.workbook;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;

import static java.io.File.createTempFile;

public class PoiWorkbookWriter implements Writable {

    private final Workbook workbook;

    public PoiWorkbookWriter(Workbook workbook) {
        this.workbook = workbook;
    }

    @Override
    public void writeTo(OutputStream stream) {
        try {
            workbook.write(stream);
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }

    @Override
    public void saveAs(File file) {
        try {
            FileOutputStream stream = new FileOutputStream(file);
            workbook.write(stream);
            stream.close();
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }

    @Override
    public void saveAsTemporaryFile(String filenameExcludingExtension) {
        String extension = workbook instanceof HSSFWorkbook ? ".xls" : ".xlsx";
        try {
            File file = createTempFile(filenameExcludingExtension, extension);
            saveAs(file);
            System.out.printf("saved %s: %s%n", extension, file.getAbsolutePath());
        } catch (IOException e) {
            throw new RuntimeException(String.format("failed to save %s: %s.%s", extension, filenameExcludingExtension, extension), e);
        }

    }
}
