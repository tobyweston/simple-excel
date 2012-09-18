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
import org.apache.poi.hssf.usermodel.HSSFHyperlink;
import org.apache.poi.ss.usermodel.Hyperlink;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;

import java.net.URL;

import static java.lang.String.format;
import static org.apache.poi.hssf.usermodel.HSSFHyperlink.LINK_URL;
import static org.apache.poi.ss.usermodel.Cell.CELL_TYPE_STRING;

/**
 * Only supports URL hyperlinks.
 */
public class HyperlinkCell extends Cell {

    private final String text;
    private final URL link;

    public HyperlinkCell(String text, URL link) {
        this(text, link, new NoStyle());
    }

    public HyperlinkCell(bad.robot.excel.Hyperlink link) {
        this(link.text(), link.value());
    }

    public HyperlinkCell(String text, URL link, Style style) {
        super(style);
        this.text = text;
        this.link = link;
    }

    @Override
    public void addTo(Row row, ColumnIndex column, Workbook workbook) {
        org.apache.poi.ss.usermodel.Cell cell = row.createCell(column.value(), CELL_TYPE_STRING);
        update(cell, workbook);
    }

    @Override
    public void update(org.apache.poi.ss.usermodel.Cell cell, Workbook workbook) {
        this.getStyle().applyTo(cell, workbook);
        cell.setCellValue(text);
        cell.setHyperlink(createHyperlink());
    }

    @Override
    public String toString() {
        return format("<a href='%s'>%s</a>", link, text);
    }

    private Hyperlink createHyperlink() {
        Hyperlink hyperlink = new HSSFHyperlink(LINK_URL);
        hyperlink.setAddress(link.toExternalForm());
        return hyperlink;
    }
}
