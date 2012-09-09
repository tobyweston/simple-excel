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
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.util.Date;

public interface WorkbookMutator {

    WorkbookMutator blankCell(Coordinate coordinate);

    WorkbookMutator replaceCell(Coordinate coordinate, String value);

    WorkbookMutator replaceCell(Coordinate coordinate, Formula formula);

    WorkbookMutator replaceCell(Coordinate coordinate, Date date);

    WorkbookMutator replaceCell(Coordinate coordinate, Double number);

    WorkbookMutator replaceCell(Coordinate coordinate, Hyperlink hyperlink);

    WorkbookMutator replaceCell(Coordinate coordinate, Boolean value);

    WorkbookMutator copyRow(Workbook workbook, Sheet worksheet, RowIndex from, RowIndex to);

    WorkbookMutator insertRowToFirstSheet(Row row, RowIndex index);

    WorkbookMutator insertRowToSheet(Row row, RowIndex index, SheetIndex sheet);

    WorkbookMutator appendRowToFirstSheet(Row row);

    WorkbookMutator appendRowToSheet(Row row, SheetIndex index);

    WorkbookMutator refreshFormulas();

}

