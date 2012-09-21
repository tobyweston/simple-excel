/*
 * Copyright (c) 2012, bad robot (london) ltd.
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
import bad.robot.excel.valuetypes.Coordinate;
import bad.robot.excel.valuetypes.RowIndex;
import bad.robot.excel.valuetypes.SheetIndex;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.util.Date;
import java.util.HashMap;
import java.util.Map;

import static bad.robot.excel.matchers.CellType.adaptPoi;
import static org.apache.poi.ss.usermodel.Row.CREATE_NULL_AS_BLANK;

/** @since 1.1 */
public class SimplifiedPoiWorkbook implements SimplifiedWorkbook {

    private final Workbook workbook;
    private final PoiWorkbookMutator mutator;

    public SimplifiedPoiWorkbook(Workbook workbook) {
        this.workbook = workbook;
        this.mutator = new PoiWorkbookMutator(workbook);
    }

    @Override
    public Cell getCellAt(Coordinate coordinate) {
        return adaptPoi(getCellForCoordinate(coordinate));
    }

    @Override
    public Row getRowAt(RowIndex rowIndex, SheetIndex sheetIndex) {
        org.apache.poi.ss.usermodel.Row row = getRowForCoordinate(rowIndex, sheetIndex);
        Map<ColumnIndex, Cell> cells = new HashMap<ColumnIndex, Cell>();
        for (org.apache.poi.ss.usermodel.Cell cell : row) {
            cells.put(null, adaptPoi(cell));
        }
        return new Row(cells);
    }

    @Override
    public WorkbookMutator blankCell(Coordinate coordinate) {
        mutator.blankCell(coordinate);
        return this;
    }

    @Override
    public WorkbookMutator replaceCell(Coordinate coordinate, Cell cell) {
        mutator.replaceCell(coordinate, cell);
        return this;
    }

    @Override
    public WorkbookMutator replaceCell(Coordinate coordinate, String value) {
        mutator.replaceCell(coordinate, value);
        return this;
    }

    @Override
    public WorkbookMutator replaceCell(Coordinate coordinate, Formula formula) {
        mutator.replaceCell(coordinate, formula);
        return this;
    }

    @Override
    public WorkbookMutator replaceCell(Coordinate coordinate, Date date) {
        mutator.replaceCell(coordinate, date);
        return this;
    }

    @Override
    public WorkbookMutator replaceCell(Coordinate coordinate, Double number) {
        mutator.replaceCell(coordinate, number);
        return this;
    }

    @Override
    public WorkbookMutator replaceCell(Coordinate coordinate, Hyperlink hyperlink) {
        mutator.replaceCell(coordinate, hyperlink);
        return this;
    }

    @Override
    public WorkbookMutator replaceCell(Coordinate coordinate, Boolean value) {
        mutator.replaceCell(coordinate, value);
        return this;
    }

    @Override
    public WorkbookMutator copyRow(Workbook workbook, Sheet worksheet, RowIndex from, RowIndex to) {
        mutator.copyRow(workbook, worksheet, from, to);
        return this;
    }

    @Override
    public WorkbookMutator insertRowToFirstSheet(Row row, RowIndex index) {
        mutator.insertRowToFirstSheet(row, index);
        return this;
    }

    @Override
    public WorkbookMutator insertRowToSheet(Row row, RowIndex index, SheetIndex sheet) {
        mutator.insertRowToSheet(row, index, sheet);
        return this;
    }

    @Override
    public WorkbookMutator appendRowToFirstSheet(Row row) {
        mutator.appendRowToFirstSheet(row);
        return this;
    }

    @Override
    public WorkbookMutator appendRowToSheet(Row row, SheetIndex index) {
        mutator.appendRowToSheet(row, index);
        return this;
    }

    @Override
    public WorkbookMutator refreshFormulas() {
        mutator.refreshFormulas();
        return this;
    }

    private org.apache.poi.ss.usermodel.Cell getCellForCoordinate(Coordinate coordinate) {
        org.apache.poi.ss.usermodel.Row row = getRowForCoordinate(coordinate.getRow(), coordinate.getSheet());
        return row.getCell(coordinate.getColumn().value(), CREATE_NULL_AS_BLANK);
    }

    private org.apache.poi.ss.usermodel.Row getRowForCoordinate(RowIndex rowIndex, SheetIndex sheetIndex) {
        Sheet sheet = workbook.getSheetAt(sheetIndex.value());
        org.apache.poi.ss.usermodel.Row row = sheet.getRow(rowIndex.value());
        if (row == null)
            row = sheet.createRow(rowIndex.value());
        return row;
    }

}
