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

import bad.robot.excel.valuetypes.ColumnIndex;

import java.util.Date;
import java.util.HashMap;
import java.util.Map;

public class StyledRowBuilder {

    private final Map<ColumnIndex, Cell> cells = new HashMap<ColumnIndex, Cell>();
    private final Style defaultStyle;

    public static StyledRowBuilder aRowWithDefaultStyle(Style defaultStyle) {
        return new StyledRowBuilder(defaultStyle);
    }

    public static StyledRowBuilder aRow() {
        return new StyledRowBuilder(new NoStyle());
    }

    private StyledRowBuilder(Style defaultStyle) {
        this.defaultStyle = defaultStyle;
    }

    public StyledRowBuilder withBlank(ColumnIndex index) {
        this.cells.put(index, new BlankCell(defaultStyle));
        return this;
    }

    public StyledRowBuilder withBlank(ColumnIndex index, Style style) {
        this.cells.put(index, new BlankCell(style));
        return this;
    }

    public StyledRowBuilder withString(ColumnIndex index, String text) {
        this.cells.put(index, new StringCell(text, defaultStyle));
        return this;
    }

    public StyledRowBuilder withString(ColumnIndex index, String text, Style style) {
        this.cells.put(index, new StringCell(text, style));
        return this;
    }

    public StyledRowBuilder withDouble(ColumnIndex index, Double value) {
        this.cells.put(index, new DoubleCell(value, defaultStyle));
        return this;
    }

    public StyledRowBuilder withDouble(ColumnIndex index, Double value, Style style) {
        this.cells.put(index, new DoubleCell(value, style));
        return this;
    }

    public StyledRowBuilder withInteger(ColumnIndex index, Integer value) {
        this.cells.put(index, new DoubleCell(new Double(Integer.toString(value)), defaultStyle));
        return this;
    }

    public StyledRowBuilder withInteger(ColumnIndex index, Integer value, Style style) {
        this.cells.put(index, new DoubleCell(new Double(Integer.toString(value)), style));
        return this;
    }

    public StyledRowBuilder withDate(ColumnIndex index, Date date) {
        this.cells.put(index, new DateCell(date, defaultStyle));
        return this;
    }

    public StyledRowBuilder withDate(ColumnIndex index, Date date, Style style) {
        this.cells.put(index, new DateCell(date, style));
        return this;
    }

    public StyledRowBuilder withFormula(ColumnIndex index, String formula) {
        this.cells.put(index, new FormulaCell(formula, defaultStyle));
        return this;
    }

    public StyledRowBuilder withFormula(ColumnIndex index, String formula, Style style) {
        this.cells.put(index, new FormulaCell(formula, style));
        return this;
    }


    public Row build() {
        return new Row(cells);
    }
}
