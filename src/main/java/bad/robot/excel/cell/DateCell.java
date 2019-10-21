/*
 * Copyright (c) 2012-2013, bad robot (london) ltd.
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *     http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */

package bad.robot.excel.cell;

import bad.robot.excel.column.ColumnIndex;
import bad.robot.excel.style.NoStyle;
import bad.robot.excel.style.Style;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;

import java.util.Date;

import static bad.robot.excel.cell.DataFormat.asDayMonthYear;
import static org.apache.poi.ss.usermodel.CellType.NUMERIC;

public class DateCell extends StyledCell {

    private final Date date;

    public DateCell(Date date) {
        this(date, new NoStyle());
    }

    public DateCell(Date date, Style style) {
        super(style);
        this.date = date;
    }

    @Override
    public void addTo(Row row, ColumnIndex column, Workbook workbook) {
        org.apache.poi.ss.usermodel.Cell cell = row.createCell(column.value(), NUMERIC);
        update(cell, workbook);
    }

    @Override
    public void update(org.apache.poi.ss.usermodel.Cell cell, Workbook workbook) {
        this.getStyle().applyTo(cell, workbook);
        if (!isCellDateFormatted(cell))
            overrideAsDateFormatting(workbook, cell);
        cell.setCellValue(date);
    }

    private void overrideAsDateFormatting(Workbook workbook, org.apache.poi.ss.usermodel.Cell cell) {
        asDayMonthYear().applyTo(cell, workbook);
    }

    @Override
    public String toString() {
        return date.toString();
    }

    private static boolean isCellDateFormatted(org.apache.poi.ss.usermodel.Cell cell) {
        return cell.getCellType() == NUMERIC && DateUtil.isCellDateFormatted(cell);
    }
}
