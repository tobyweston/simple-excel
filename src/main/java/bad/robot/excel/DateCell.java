/*
 * Copyright (c) 2012-2013, bad robot (london) ltd.
 *
 * Licensed to the Apache Software Foundation (ASF) under one
 * or more contributor license agreements.  See the NOTICE file
 * distributed with this work for additional information
 * regarding copyright ownership.  The ASF licenses this file
 * to you under the Apache License, Version 2.0 (the
 * "License"); you may not use this file except in compliance
 * with the License.  You may obtain a copy of the License at
 *
 * http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing,
 * software distributed under the License is distributed on an
 * "AS IS" BASIS, WITHOUT WARRANTIES OR CONDITIONS OF ANY
 * KIND, either express or implied.  See the License for the
 * specific language governing permissions and limitations
 * under the License.
 */

package bad.robot.excel;

import bad.robot.excel.valuetypes.ColumnIndex;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;

import java.util.Date;

import static bad.robot.excel.valuetypes.DataFormat.asDayMonthYear;
import static org.apache.poi.ss.usermodel.Cell.CELL_TYPE_NUMERIC;

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
        org.apache.poi.ss.usermodel.Cell cell = row.createCell(column.value(), CELL_TYPE_NUMERIC);
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
        return cell.getCellType() == CELL_TYPE_NUMERIC && HSSFDateUtil.isCellDateFormatted(cell);
    }
}
