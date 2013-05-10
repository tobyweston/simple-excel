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

import bad.robot.excel.valuetypes.Coordinate;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;

public class WorkbookResource {

    public static Workbook getWorkbook(String file) throws IOException {
        InputStream stream = PoiWorkbookMutatorTest.class.getResourceAsStream(file);
        if (stream == null)
            throw new FileNotFoundException(file);
        return new HSSFWorkbook(stream);
    }

    public static Sheet firstSheetOf(String file) throws IOException {
        return getWorkbook(file).getSheetAt(0);
    }

    public static org.apache.poi.ss.usermodel.Row firstRowOf(String file) throws IOException {
        return firstSheetOf(file).getRow(0);
    }

    public static org.apache.poi.ss.usermodel.Row secondRowOf(String file) throws IOException {
        return firstSheetOf(file).getRow(1);
    }

    public static org.apache.poi.ss.usermodel.Row thirdRowOf(String file) throws IOException {
        return firstSheetOf(file).getRow(2);
    }

    public static org.apache.poi.ss.usermodel.Row firstRowOf(Sheet sheet) {
        return sheet.getRow(0);
    }

    public static org.apache.poi.ss.usermodel.Cell getCellForCoordinate(Coordinate coordinate, Workbook workbook) throws IOException {
        org.apache.poi.ss.usermodel.Row row = getRowForCoordinate(coordinate, workbook);
        return row.getCell(coordinate.getColumn().value());
    }

    public static org.apache.poi.ss.usermodel.Row getRowForCoordinate(Coordinate coordinate, Workbook workbook) throws IOException {
        Sheet sheet = workbook.getSheetAt(coordinate.getSheet().value());
        org.apache.poi.ss.usermodel.Row row = sheet.getRow(coordinate.getRow().value());
        if (row == null)
            throw new IllegalStateException("expected to find a row");
        return row;
    }

    public static String getCellDataFormatAtCoordinate(Coordinate coordinate, Workbook workbook) throws IOException {
        return getCellForCoordinate(coordinate, workbook).getCellStyle().getDataFormatString();
    }
}
