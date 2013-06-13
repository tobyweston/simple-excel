package bad.robot.excel.workbook;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public enum WorkbookType {
    XLS("Excel Binary File Format (pre 2007)") {
        @Override
        public Workbook create() {
            HSSFWorkbook workbook = new HSSFWorkbook();
            workbook.createSheet("Sheet1");
            return workbook;
        }
    },
    XML("Office Open XML") {
        @Override
        public Workbook create() {
            XSSFWorkbook workbook = new XSSFWorkbook();
            workbook.createSheet("Sheet1");
            return workbook;
        }
    };

    private final String description;

    WorkbookType(String description) {
        this.description = description;
    }

    public abstract Workbook create();

    public String getDescription() {
        return description;
    }
}
