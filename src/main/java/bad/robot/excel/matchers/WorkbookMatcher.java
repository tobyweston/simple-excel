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

import org.apache.poi.ss.usermodel.Workbook;
import org.hamcrest.Description;
import org.hamcrest.Matcher;
import org.hamcrest.TypeSafeDiagnosingMatcher;

import static bad.robot.excel.matchers.SheetMatcher.hasSameSheetsAs;

public class WorkbookMatcher extends TypeSafeDiagnosingMatcher<Workbook> {

    private final Workbook expected;

    public static Matcher<Workbook> sameWorkbook(Workbook expectedWorkbook) {
        return new WorkbookMatcher(expectedWorkbook);
    }

    private WorkbookMatcher(Workbook expected) {
        this.expected = expected;
    }

    @Override
    protected boolean matchesSafely(Workbook actual, Description mismatch) {
        if (!hasSameSheetsAs(expected).matchesSafely(actual, mismatch))
            return false;

        if (!new SheetsMatcher(expected).matchesSafely(actual, mismatch))
            return false;

        return true;
    }


    @Override
    public void describeTo(Description description) {
        description.appendText("entire workbook to be equal");
    }

}
