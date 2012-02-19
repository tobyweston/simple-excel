package bad.robot.excel;

import java.io.IOException;
import java.io.InputStream;

public interface ExcelWorkbook {

    InputStream getInputStream() throws IOException;

}
