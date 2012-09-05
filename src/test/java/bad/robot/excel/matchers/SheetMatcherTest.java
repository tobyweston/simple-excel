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

import org.apache.poi.ss.usermodel.Workbook;
import org.hamcrest.Description;
import org.hamcrest.StringDescription;
import org.junit.Before;
import org.junit.Test;

import java.io.IOException;

import static bad.robot.excel.WorkbookResource.getWorkbook;
import static bad.robot.excel.matchers.SheetMatcher.hasSameSheetsAs;
import static org.hamcrest.MatcherAssert.assertThat;
import static org.hamcrest.Matchers.*;

public class SheetMatcherTest {

    private final Description description = new StringDescription();

    private Workbook workbookWithOneNamedSheet;
    private Workbook workbookWithManyNamedSheets;
    private Workbook workbookWithMultipleAlternativelyNamedSheets;

    @Before
    public void loadWorkbooks() throws IOException {
        workbookWithOneNamedSheet = getWorkbook("workbookWithOneSheet.xls");
        workbookWithManyNamedSheets = getWorkbook("workbookWithMultipleNamedSheets.xls");
        workbookWithMultipleAlternativelyNamedSheets = getWorkbook("workbookWithMultipleAlternativelyNamedSheets.xls");
    }

    @Test
    public void exampleUsages() {
        assertThat(workbookWithOneNamedSheet, not(hasSameSheetsAs(workbookWithManyNamedSheets)));
    }

    @Test
    public void matches() {
        assertThat(hasSameSheetsAs(workbookWithManyNamedSheets).matches(workbookWithManyNamedSheets), is(true));
    }

    @Test
    public void doesNotMatch() {
        assertThat(hasSameSheetsAs(workbookWithManyNamedSheets).matches(workbookWithOneNamedSheet), is(false));
    }

    @Test
    public void description() {
        hasSameSheetsAs(workbookWithManyNamedSheets).describeTo(description);
        assertThat(description.toString(), is("<4> sheet(s) named \"Sheet1\", \"Another Sheet\", \"Yet Another Sheet\", \"Sheet5\""));
    }

    @Test
    public void mismatchOnNumberOfSheets() {
        hasSameSheetsAs(workbookWithManyNamedSheets).matchesSafely(workbookWithOneNamedSheet, description);
        assertThat(description.toString(), allOf(
            containsString("got <1> sheet(s) expected <4>,"),
            containsString("sheet(s) \"Another Sheet\", \"Yet Another Sheet\", \"Sheet5\" were missing")
        ));
    }

    @Test
    public void mismatchOnNameOfSheets() {
        hasSameSheetsAs(workbookWithManyNamedSheets).matchesSafely(workbookWithMultipleAlternativelyNamedSheets, description);
        assertThat(description.toString(), is("sheet(s) \"Yet Another Sheet\" was missing"));
    }
}
