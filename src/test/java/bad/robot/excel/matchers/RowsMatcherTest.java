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

package bad.robot.excel.matchers;

import org.apache.poi.ss.usermodel.Sheet;
import org.hamcrest.StringDescription;
import org.junit.Before;
import org.junit.Test;

import java.io.IOException;

import static bad.robot.excel.WorkbookResource.firstSheetOf;
import static bad.robot.excel.matchers.RowsMatcher.hasSameRowsAs;
import static org.hamcrest.MatcherAssert.assertThat;
import static org.hamcrest.Matchers.*;

public class RowsMatcherTest {

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
        assertThat(sheetWithDifferingValues, not(hasSameRowsAs(sheetWithThreeRows)));
        assertThat(sheetWithTwoRows, not(hasSameRowsAs(sheetWithThreeRows)));
        assertThat(sheetWithThreeRows, hasSameRowsAs(sheetWithTwoRows));
    }

    @Test
    public void matches() throws IOException {
        assertThat(hasSameRowsAs(sheetWithThreeRows).matches(sheetWithThreeRows), is(true));
    }

    @Test
    public void doesNotMatch() {
        assertThat(hasSameRowsAs(sheetWithThreeRows).matches(sheetWithDifferingValues), is(false));
    }

    @Test
    public void description() throws IOException {
        hasSameRowsAs(sheetWithThreeRows).describeTo(description);
        assertThat(description.toString(), containsString("equality on all rows in \"Sheet1\""));
    }

    @Test
    public void mismatch() throws IOException {
        hasSameRowsAs(sheetWithThreeRows).matchesSafely(sheetWithDifferingValues, description);
        assertThat(description.toString(), containsString("cell at \"A2\" contained <\"XXX\"> expected <\"Row 2\">"));
    }

    @Test
    public void mismatchOnMissingRow() {
        hasSameRowsAs(sheetWithThreeRows).matchesSafely(sheetWithTwoRows, description);
        assertThat(description.toString(), is("row <3> is missing"));
    }
}
