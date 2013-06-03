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

package bad.robot.excel.row;

import bad.robot.excel.cell.*;
import bad.robot.excel.column.ColumnIndex;
import bad.robot.excel.column.ExcelColumnIndex;
import bad.robot.excel.style.Style;

import java.util.Date;
import java.util.HashMap;
import java.util.Map;

import static bad.robot.excel.column.ColumnIndex.column;


public class NullSkippingRowBuilder implements RowBuilder {

    private final Map<ColumnIndex, Cell> cells;
    private final Style defaultStyle;

    public NullSkippingRowBuilder(int initialCapacity, Style defaultStyle) {
        this.cells = new HashMap<ColumnIndex, Cell>(initialCapacity);
        this.defaultStyle = defaultStyle;
        for (int i = 0; i < initialCapacity; i++)
            cells.put(column(ExcelColumnIndex.from(i)), new BlankCell(defaultStyle));
    }

    @Override
    public RowBuilder withBlank(ColumnIndex index) {
        this.cells.put(index, new BlankCell(defaultStyle));
        return this;
    }

    @Override
    public NullSkippingRowBuilder withString(ColumnIndex index, String text) {
        if (text != null)
            this.cells.put(index, new StringCell(text, defaultStyle));
        return this;
    }

    @Override
    public NullSkippingRowBuilder withDouble(ColumnIndex index, Double value) {
        if (value != null)
            this.cells.put(index, new DoubleCell(value, defaultStyle));
        return this;
    }

    @Override
    public NullSkippingRowBuilder withInteger(ColumnIndex index, Integer value) {
        if (value != null)
            this.cells.put(index, new DoubleCell(new Double(Integer.toString(value)), defaultStyle));
        return this;
    }

    @Override
    public NullSkippingRowBuilder withDate(ColumnIndex index, Date date) {
        if (date != null)
            this.cells.put(index, new DateCell(date, defaultStyle));
        return this;
    }

    @Override
    public RowBuilder withFormula(ColumnIndex index, String formula) {
        if (formula != null)
            this.cells.put(index, new StringCell(formula, defaultStyle));
        return this;
    }

    @Override
    public Row build() {
        return new Row(cells);
    }
}
