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

package bad.robot.excel.workbook;

import bad.robot.excel.cell.*;
import bad.robot.excel.row.CopyRow;
import bad.robot.excel.row.Row;
import bad.robot.excel.row.RowIndex;
import bad.robot.excel.sheet.Coordinate;
import bad.robot.excel.sheet.SheetIndex;
import org.apache.poi.hssf.usermodel.HSSFFormulaEvaluator;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.util.Date;

import static org.apache.poi.ss.usermodel.Row.CREATE_NULL_AS_BLANK;


public class PoiWorkbook implements Editable {

    private final org.apache.poi.ss.usermodel.Workbook workbook;

    public PoiWorkbook(WorkbookType type) {
        workbook = type.create();
    }

    public PoiWorkbook(InputStream stream) throws IOException {
        workbook = new PoiWorkbookReader().read(stream);
    }

    public PoiWorkbook(File file) throws IOException {
        workbook = new PoiWorkbookReader().read(file);
    }

    public PoiWorkbook(org.apache.poi.ss.usermodel.Workbook workbook) {
        if (workbook == null)
            throw new IllegalArgumentException();
        this.workbook = workbook;
    }

    @Override
    public PoiWorkbook blankCell(Coordinate coordinate) {
        new BlankCell().update(getCellForCoordinate(coordinate), workbook);
        return this;
    }

    @Override
    public PoiWorkbook replaceCell(Coordinate coordinate, Cell cell) {
        cell.update(getCellForCoordinate(coordinate), workbook);
        return this;
    }

    @Override
    public PoiWorkbook replaceCell(Coordinate coordinate, String text) {
        return replaceCell(coordinate, new StringCell(text));
    }

    @Override
    public PoiWorkbook replaceCell(Coordinate coordinate, Formula formula) {
        return replaceCell(coordinate, new FormulaCell(formula));
    }

    @Override
    public PoiWorkbook replaceCell(Coordinate coordinate, Date date) {
        return replaceCell(coordinate, new DateCell(date));
    }

    @Override
    public PoiWorkbook replaceCell(Coordinate coordinate, Double number) {
        return replaceCell(coordinate, new DoubleCell(number));
    }

    @Override
    public PoiWorkbook replaceCell(Coordinate coordinate, Hyperlink hyperlink) {
        return replaceCell(coordinate, new HyperlinkCell(hyperlink));
    }

    @Override
    public PoiWorkbook replaceCell(Coordinate coordinate, Boolean value) {
        return replaceCell(coordinate, new BooleanCell(value));
    }

    @Override
    public PoiWorkbook copyRow(org.apache.poi.ss.usermodel.Workbook workbook, Sheet worksheet, RowIndex from, RowIndex to) {
        CopyRow.copyRow(workbook, worksheet, from, to);
        return this;
    }

    @Override
    public PoiWorkbook insertRowToFirstSheet(Row row, RowIndex index) {
        row.insertAt(workbook, SheetIndex.sheet(1), index);
        return this;
    }

    @Override
    public PoiWorkbook insertRowToSheet(Row row, RowIndex index, SheetIndex sheet) {
        row.insertAt(workbook, sheet, index);
        return this;
    }

    @Override
    public PoiWorkbook appendRowToFirstSheet(Row row) {
        row.appendTo(workbook, SheetIndex.sheet(1));
        return this;
    }

    @Override
    public PoiWorkbook appendRowToSheet(Row row, SheetIndex index) {
        row.appendTo(workbook, index);
        return this;
    }

    @Override
    public PoiWorkbook refreshFormulas() {
        HSSFFormulaEvaluator.evaluateAllFormulaCells(workbook);
        return this;
    }

    private org.apache.poi.ss.usermodel.Cell getCellForCoordinate(Coordinate coordinate) {
        org.apache.poi.ss.usermodel.Row row = getRowForCoordinate(coordinate);
        return row.getCell(coordinate.getColumn().value(), CREATE_NULL_AS_BLANK);
    }

    private org.apache.poi.ss.usermodel.Row getRowForCoordinate(Coordinate coordinate) {
        Sheet sheet = workbook.getSheetAt(coordinate.getSheet().value());
        org.apache.poi.ss.usermodel.Row row = sheet.getRow(coordinate.getRow().value());
        if (row == null)
            row = sheet.createRow(coordinate.getRow().value());
        return row;
    }

    /**
     * Accessor to expose the underlying Poi workbook in case I've missed something and clients need to work
     * with the raw workbook directly.
     *
     * <b>Try not to use this</b>
     */
    public Workbook workbook() {
        return workbook;
    }
}
