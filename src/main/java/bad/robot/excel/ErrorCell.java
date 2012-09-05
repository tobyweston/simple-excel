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
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;

import static java.lang.String.format;
import static org.apache.poi.ss.usermodel.Cell.CELL_TYPE_ERROR;

public class ErrorCell extends Cell {

    private final Byte value;

    public ErrorCell(Byte text) {
        this(text, new NoStyle());
    }

    public ErrorCell(Byte text, Style style) {
        super(style);
        this.value = text;
    }

    @Override
    public void addTo(Row row, ColumnIndex column, Workbook workbook) {
        org.apache.poi.ss.usermodel.Cell cell = row.createCell(column.value(), CELL_TYPE_ERROR);
        this.getStyle().applyTo(cell, workbook);
        cell.setCellErrorValue(value);
    }

    @Override
    public String toString() {
        return format("Error:%s", value);
    }
}