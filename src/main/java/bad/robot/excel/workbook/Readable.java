package bad.robot.excel.workbook;

import org.apache.poi.ss.usermodel.Workbook;

import java.io.File;
import java.io.IOException;
import java.io.InputStream;

public interface Readable<T> {

    Workbook read(File file) throws IOException;

    Workbook read(InputStream stream) throws IOException;
}
