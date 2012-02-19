package bad.robot.excel;

import org.apache.poi.ss.usermodel.Workbook;

public interface Style {

    void applyTo(org.apache.poi.ss.usermodel.Cell cell, Workbook template);

}
