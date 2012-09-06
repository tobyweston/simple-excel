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

import bad.robot.excel.valuetypes.Coordinate;
import bad.robot.excel.valuetypes.RowIndex;
import bad.robot.excel.valuetypes.SheetIndex;
import org.apache.poi.hssf.usermodel.HSSFFormulaEvaluator;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.util.Date;

import static org.apache.poi.ss.usermodel.Row.CREATE_NULL_AS_BLANK;


public class PoiWorkbookMutator implements WorkbookMutator {

    private final Workbook workbook;

    public PoiWorkbookMutator(Workbook workbook) {
        if (workbook == null)
            throw new IllegalArgumentException();
        this.workbook = workbook;
    }

    @Override
    public WorkbookMutator replaceCell(Coordinate coordinate, String text) {
        new StringCell(text).update(getCellForCoordinate(coordinate), workbook);
        return this;
    }

    @Override
    public WorkbookMutator replaceCell(Coordinate coordinate, Formula formula) {
        new FormulaCell(formula).update(getCellForCoordinate(coordinate), workbook);
        return this;
    }

    @Override
    public WorkbookMutator replaceCell(Coordinate coordinate, Date date) {
        new DateCell(date).update(getCellForCoordinate(coordinate), workbook);
        return this;
    }

    @Override
    public WorkbookMutator replaceCell(Coordinate coordinate, Double value) {
        new DoubleCell(value).update(getCellForCoordinate(coordinate), workbook);
        return this;
    }

    @Override
    public WorkbookMutator replaceCell(Coordinate coordinate, Hyperlink link) {
        new HyperlinkCell(link).update(getCellForCoordinate(coordinate), workbook);
        return this;
    }

    // TODO add more types to "replace" methods

    @Override
    public WorkbookMutator copyRow(Workbook workbook, Sheet worksheet, RowIndex from, RowIndex to) {
        CopyRow.copyRow(workbook, worksheet, from, to);
        return this;
    }

    @Override
    public WorkbookMutator insertRowToFirstSheet(Row row, RowIndex index) {
        row.insertAt(workbook, SheetIndex.sheet(1), index);
        return this;
    }

    @Override
    public WorkbookMutator insertRowToSheet(Row row, RowIndex index, SheetIndex sheet) {
        row.insertAt(workbook, sheet, index);
        return this;
    }

    @Override
    public WorkbookMutator appendRowToFirstSheet(Row row) {
        row.appendTo(workbook, SheetIndex.sheet(1));
        return this;
    }

    @Override
    public WorkbookMutator appendRowToSheet(Row row, SheetIndex index) {
        row.appendTo(workbook, index);
        return this;
    }

    @Override
    public WorkbookMutator refreshFormulas() {
        HSSFFormulaEvaluator.evaluateAllFormulaCells((HSSFWorkbook) workbook);
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
}