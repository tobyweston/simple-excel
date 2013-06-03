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

package bad.robot.excel.cell;

import bad.robot.excel.AbstractValueType;
import bad.robot.excel.style.Style;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Workbook;

import static bad.robot.excel.style.ClonedStyleFactory.newStyleBasedOn;

public class DataFormat extends AbstractValueType<String> implements Style {

    public static DataFormat asDayMonthYear() {
        return dataFormat("dd-MMM-yyyy");
    }

    public static DataFormat asTwoDecimalPlacesNumber() {
        return dataFormat("#,##0.00");
    }

    public static DataFormat dataFormat(String value) {
        return new DataFormat(value);
    }

    private DataFormat(String value) {
        super(value);
    }

    @Override
    public void applyTo(Cell cell, Workbook workbook) {
        updateDataFormat(cell, workbook);
    }

    private void updateDataFormat(Cell cell, Workbook workbook) {
        CellStyle style = newStyleBasedOn(cell).create(workbook);
        style.setDataFormat(workbook.createDataFormat().getFormat(value()));
        cell.setCellStyle(style);
    }

}
