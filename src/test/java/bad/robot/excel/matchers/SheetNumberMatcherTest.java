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

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.hamcrest.Description;
import org.hamcrest.StringDescription;
import org.junit.Before;
import org.junit.Test;

import java.io.IOException;

import static bad.robot.excel.WorkbookResource.getWorkbook;
import static bad.robot.excel.matchers.SheetNumberMatcher.hasSameNumberOfSheetsAs;
import static org.hamcrest.MatcherAssert.assertThat;
import static org.hamcrest.Matchers.is;
import static org.hamcrest.Matchers.not;

// Using Assert (from JUnit)
//java.lang.AssertionError:
//        Expected: <2> sheet(s)
//        got: <org.apache.poi.hssf.usermodel.HSSFWorkbook@3b6f0be8>

// Using MatcherAssert:
//java.lang.AssertionError:
//        Expected: <2> sheet(s)
//        but: got <1> sheet(s)
//        at org.hamcrest.MatcherAssert.assertThat(MatcherAssert.java:20)

public class SheetNumberMatcherTest {

    private HSSFWorkbook workbookWithOneSheet;
    private HSSFWorkbook workbookWithTwoSheets;

    @Before
    public void loadWorkbooks() throws IOException {
        workbookWithOneSheet = getWorkbook("workbookWithOneSheet.xls");
        workbookWithTwoSheets = getWorkbook("workbookWithTwoSheets.xls");
    }

    @Test
    public void exampleUsages() throws Exception {
        assertThat(workbookWithOneSheet, hasSameNumberOfSheetsAs(workbookWithOneSheet));
        assertThat(workbookWithOneSheet, not(hasSameNumberOfSheetsAs(workbookWithTwoSheets)));
    }

    @Test
    public void matches() {
        assertThat(hasSameNumberOfSheetsAs(workbookWithOneSheet).matches(workbookWithOneSheet), is(true));
    }

    @Test
    public void doesNotMatch() {
        assertThat(hasSameNumberOfSheetsAs(workbookWithOneSheet).matches(workbookWithTwoSheets), is(false));
    }

    @Test
    public void description() {
        Description description = new StringDescription();
        hasSameNumberOfSheetsAs(workbookWithOneSheet).describeTo(description);
        assertThat(description.toString(), is("<1> sheet(s)"));
    }

    @Test
    public void mismatch() {
        Description description = new StringDescription();
        hasSameNumberOfSheetsAs(workbookWithOneSheet).matchesSafely(workbookWithTwoSheets, description);
        assertThat(description.toString(), is("got <2> sheet(s)"));
    }
}
