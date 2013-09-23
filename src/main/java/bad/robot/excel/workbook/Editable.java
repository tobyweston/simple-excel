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

package bad.robot.excel.workbook;

import bad.robot.excel.cell.Cell;
import bad.robot.excel.cell.Formula;
import bad.robot.excel.cell.Hyperlink;
import bad.robot.excel.row.Row;
import bad.robot.excel.row.RowIndex;
import bad.robot.excel.sheet.Coordinate;
import bad.robot.excel.sheet.SheetIndex;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.util.Date;

public interface Editable {

    Editable blankCell(Coordinate coordinate);

    Editable replaceCell(Coordinate coordinate, Cell cell);

    Editable replaceCell(Coordinate coordinate, String value);

    Editable replaceCell(Coordinate coordinate, Formula formula);

    Editable replaceCell(Coordinate coordinate, Date date);

    Editable replaceCell(Coordinate coordinate, Double number);

    Editable replaceCell(Coordinate coordinate, Hyperlink hyperlink);

    Editable replaceCell(Coordinate coordinate, Boolean value);

    Editable copyRow(Workbook workbook, Sheet worksheet, RowIndex from, RowIndex to);

    Editable insertSheet();

    Editable insertSheet(String name);

    Editable insertRowToFirstSheet(Row row, RowIndex index);

    Editable insertRowToSheet(Row row, RowIndex index, SheetIndex sheet);

    Editable appendRowToFirstSheet(Row row);

    Editable appendRowToSheet(Row row, SheetIndex index);

    Editable refreshFormulas();

}

