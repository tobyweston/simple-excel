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

package bad.robot.excel.sheet;

import bad.robot.excel.column.ColumnIndex;
import bad.robot.excel.column.ExcelColumnIndex;
import bad.robot.excel.row.RowIndex;

import static bad.robot.excel.sheet.SheetIndex.sheet;

public class Coordinate {

    private final ColumnIndex column;
    private final RowIndex row;
    private final SheetIndex sheet;

    public static Coordinate coordinate(ExcelColumnIndex column, Integer row) {
        return new Coordinate(ColumnIndex.column(column), RowIndex.row(row), sheet(1));
    }

    public static Coordinate coordinate(ExcelColumnIndex column, Integer row, SheetIndex sheet) {
        return new Coordinate(ColumnIndex.column(column), RowIndex.row(row), sheet);
    }

    public static Coordinate coordinate(ColumnIndex column, RowIndex row) {
        return new Coordinate(column, row, sheet(1));
    }

    public static Coordinate coordinate(ColumnIndex column, RowIndex row, SheetIndex sheet) {
        return new Coordinate(column, row, sheet);
    }

    private Coordinate(ColumnIndex column, RowIndex row, SheetIndex sheet) {
        this.column = column;
        this.row = row;
        this.sheet = sheet;
    }

    public ColumnIndex getColumn() {
        return column;
    }

    public RowIndex getRow() {
        return row;
    }

    public SheetIndex getSheet() {
        return sheet;
    }
}