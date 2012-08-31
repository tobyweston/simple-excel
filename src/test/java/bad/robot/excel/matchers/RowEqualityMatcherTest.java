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

package bad.robot.excel.matchers;

import org.apache.poi.ss.usermodel.Sheet;
import org.hamcrest.StringDescription;
import org.junit.Before;
import org.junit.Test;

import java.io.IOException;

import static bad.robot.excel.WorkbookResource.firstSheetOf;
import static bad.robot.excel.matchers.RowEqualityMatcher.rowsEqual;
import static org.hamcrest.MatcherAssert.assertThat;
import static org.hamcrest.Matchers.is;
import static org.hamcrest.Matchers.not;

public class RowEqualityMatcherTest {

    private final StringDescription description = new StringDescription();

    private Sheet sheetWithDifferingValues;
    private Sheet sheetWithThreeRows;
    private Sheet sheetWithTwoRows;

    @Before
    public void loadWorkbooks() throws IOException {
        sheetWithThreeRows = firstSheetOf("sheetWithThreeRows.xls");
        sheetWithDifferingValues = firstSheetOf("sheetWithThreeRowsWithAlternativeValues.xls");
        sheetWithTwoRows = firstSheetOf("sheetWithTwoRows.xls");
    }

    @Test
    public void exampleUsage() throws IOException {
        assertThat(sheetWithDifferingValues, not(rowsEqual(sheetWithThreeRows)));
        assertThat(sheetWithTwoRows, not(rowsEqual(sheetWithThreeRows)));
        assertThat(sheetWithThreeRows, rowsEqual(sheetWithTwoRows));
    }

    @Test
    public void matches() throws IOException {
        assertThat(rowsEqual(sheetWithThreeRows).matches(sheetWithThreeRows), is(true));
    }

    @Test
    public void doesNotMatch() {
        assertThat(rowsEqual(sheetWithThreeRows).matches(sheetWithDifferingValues), is(false));
    }

    @Test
    public void description() throws IOException {
        rowsEqual(sheetWithThreeRows).describeTo(description);
        assertThat(description.toString(), is("equality on all rows in \"Sheet1\""));
    }

    @Test
    public void mismatch() throws IOException {
        rowsEqual(sheetWithThreeRows).matchesSafely(sheetWithDifferingValues, description);
        assertThat(description.toString(), is("cell at \"A2\" contained <\"XXX\"> expected <\"Row 2\">"));
    }

    @Test
    public void mismatchOn() {
        rowsEqual(sheetWithThreeRows).matchesSafely(sheetWithTwoRows, description);
        assertThat(description.toString(), is("row 3 is missing"));
    }
}
