package bad.robot.excel;

import org.apache.poi.ss.usermodel.Workbook;

public class NoStyle implements Style {

    @Override
    public void applyTo(org.apache.poi.ss.usermodel.Cell cell, Workbook template) {
        // Do nothing
    }
}
