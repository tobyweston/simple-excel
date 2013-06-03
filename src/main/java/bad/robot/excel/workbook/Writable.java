package bad.robot.excel.workbook;

import java.io.File;
import java.io.OutputStream;

public interface Writable {

    void writeTo(OutputStream stream);

    void saveAs(File file);

    void saveAsTemporaryFile(String filenameExcludingExtension);
}
