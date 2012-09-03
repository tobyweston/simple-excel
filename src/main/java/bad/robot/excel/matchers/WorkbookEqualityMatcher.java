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
import org.apache.poi.ss.usermodel.Workbook;
import org.hamcrest.Description;
import org.hamcrest.Matcher;
import org.hamcrest.TypeSafeMatcher;

import static bad.robot.excel.matchers.RowNumberMatcher.hasSameNumberOfRowAs;
import static bad.robot.excel.matchers.RowsMatcher.hasSameRowsAs;
import static bad.robot.excel.matchers.SheetMatcher.hasSameSheetsAs;

public class WorkbookEqualityMatcher extends TypeSafeMatcher<Workbook> {

    private final Workbook expectedWorkbook;

    WorkbookEqualityMatcher(Workbook expectedWorkbook) {
        this.expectedWorkbook = expectedWorkbook;
    }

    public static Matcher<Workbook> sameWorkBook(Workbook expectedWorkbook) {
        return new WorkbookEqualityMatcher(expectedWorkbook);
    }

    @Override
    public boolean matchesSafely(Workbook actual) {
        if (!hasSameSheetsAs(expectedWorkbook).matches(actual))
            return false;

        for (int a = 0; a < actual.getNumberOfSheets(); a++) {
            Sheet actualSheet = actual.getSheetAt(a);
            Sheet expectedSheet = expectedWorkbook.getSheetAt(a);

            if (!hasSameNumberOfRowAs(expectedSheet).matches(actualSheet))
                return false;

            if (!hasSameRowsAs(expectedSheet).matches(actualSheet))
                return false;
        }

        return true;
    }

    @Override
    public void describeTo(Description description) {
        description.appendText("entire workbook to be equal");
    }


}
