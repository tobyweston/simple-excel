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
import bad.robot.excel.valuetypes.RowIndex;
import bad.robot.excel.valuetypes.SheetIndex;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.util.HashMap;
import java.util.Map;

public class Row {

    private static final int shiftDownAmount = 1;

    private final Map<ColumnIndex, Cell> cells = new HashMap<ColumnIndex, Cell>();

    public Row(Map<ColumnIndex, Cell> cells) {
        this.cells.putAll(cells);
    }

    public void insertAt(Workbook workbook, SheetIndex sheetIndex, RowIndex rowIndex) {
        Sheet sheet = workbook.getSheetAt(sheetIndex.value());
        sheet.shiftRows(rowIndex.value(), sheet.getLastRowNum(), shiftDownAmount);
        org.apache.poi.ss.usermodel.Row row = sheet.createRow(rowIndex.value());
        copyCellsTo(row, workbook);
    }

    public void appendTo(Workbook workbook, SheetIndex index) {
        Sheet sheet = workbook.getSheetAt(index.value());
        org.apache.poi.ss.usermodel.Row row = createRow(sheet);
        copyCellsTo(row, workbook);
    }

    private void copyCellsTo(org.apache.poi.ss.usermodel.Row row, Workbook workbook) {
        for (ColumnIndex index : cells.keySet()) {
            Cell cellToInsert = cells.get(index);
            cellToInsert.addTo(row, index, workbook);
        }
    }

    private static org.apache.poi.ss.usermodel.Row createRow(Sheet sheet) {
        if (sheet.getPhysicalNumberOfRows() == 0)
            return sheet.createRow(0);
        return sheet.createRow(sheet.getLastRowNum() + 1);
    }


}
