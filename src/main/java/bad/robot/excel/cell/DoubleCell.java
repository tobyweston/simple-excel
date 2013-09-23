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
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;

import static java.lang.String.format;
import static org.apache.poi.ss.usermodel.Cell.CELL_TYPE_NUMERIC;

public class DoubleCell extends StyledCell {

    private final Double number;

    public DoubleCell(Double number) {
        this(number, new NoStyle());
    }

    public DoubleCell(Double number, Style style) {
        super(style);
        this.number = number;
    }

    @Override
    public void addTo(Row row, ColumnIndex column, Workbook workbook) {
        org.apache.poi.ss.usermodel.Cell cell = row.createCell(column.value(), CELL_TYPE_NUMERIC);
        update(cell, workbook);
    }

    @Override
    public void update(org.apache.poi.ss.usermodel.Cell cell, Workbook workbook) {
        this.getStyle().applyTo(cell, workbook);
        cell.setCellValue(number);
    }

    @Override
    public String toString() {
        return format("%sD", number.toString());
    }
}
