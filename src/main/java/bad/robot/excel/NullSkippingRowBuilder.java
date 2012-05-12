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

import bad.robot.excel.valuetypes.CellIndex;
import bad.robot.excel.valuetypes.ExcelColumnIndex;

import java.util.Date;
import java.util.HashMap;
import java.util.Map;

import static bad.robot.excel.valuetypes.CellIndex.cellIndex;

public class NullSkippingRowBuilder {

    private final Map<CellIndex, Cell> cells;
    private final Style defaultStyle;

    public NullSkippingRowBuilder(int initialCapacity, Style defaultStyle) {
        cells = new HashMap<CellIndex, Cell>(initialCapacity);
        this.defaultStyle = defaultStyle;
        for (int i = 0; i < initialCapacity; i++)
            cells.put(cellIndex(ExcelColumnIndex.values()[i]), new BlankCell(defaultStyle));
    }

    public NullSkippingRowBuilder withString(CellIndex index, String text) {
        if (text != null)
            this.cells.put(index, new StringCell(text, defaultStyle));
        return this;
    }

    public NullSkippingRowBuilder withDouble(CellIndex index, Double value) {
        if (value != null)
            this.cells.put(index, new DoubleCell(value, defaultStyle));
        return this;
    }

    public NullSkippingRowBuilder withDouble(CellIndex index, Double value, Style style) {
        if (value != null)
            this.cells.put(index, new DoubleCell(value, style));
        return this;
    }

    public NullSkippingRowBuilder withInteger(CellIndex index, Integer value) {
        if (value != null)
            this.cells.put(index, new DoubleCell(new Double(Integer.toString(value)), defaultStyle));
        return this;
    }

    public NullSkippingRowBuilder withDate(CellIndex index, Date date) {
        if (date != null)
            this.cells.put(index, new DateCell(date, defaultStyle));
        return this;
    }

    public NullSkippingRowBuilder withDate(CellIndex index, Date date, Style style) {
        if (date != null)
            this.cells.put(index, new DateCell(date, style));
        return this;
    }

    public Row build() {
        return new Row(cells);
    }
}
