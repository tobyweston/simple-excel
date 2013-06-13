package bad.robot.excel.workbook;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;

public class PoiWorkbookReader implements Readable<Workbook> {

    @Override
    public Workbook read(File file) throws IOException {
        InputStream stream = new FileInputStream(file);
        return read(stream);
    }

    @Override
    public Workbook read(InputStream stream) throws IOException {
        if (stream == null)
            throw new IllegalArgumentException("stream was null, could not load the workbook");
        try {
            return WorkbookFactory.create(stream);
        } catch (InvalidFormatException e) {
            throw new IOException(e);
        } finally {
            stream.close();
        }
    }

}
