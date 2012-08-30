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

import org.junit.Test;

import static bad.robot.excel.WorkbookResource.firstSheetOf;
import static bad.robot.excel.matchers.RowNumberMatcher.hasSameNumberOfRowAs;
import static org.hamcrest.MatcherAssert.assertThat;
import static org.hamcrest.Matchers.not;

// Using Assert (from JUnit)
//java.lang.AssertionError:
//        Expected: <3> row(s) in sheet "Sheet1"
//        got: <org.apache.poi.hssf.usermodel.HSSFSheet@47d62270>

// Using MatcherAssert:
//java.lang.AssertionError:
//        Expected: <3> row(s) in sheet "Sheet1"
//        but: got <2> row(s) in sheet "Sheet1"
//        at org.hamcrest.MatcherAssert.assertThat(MatcherAssert.java:20)

public class RowNumberMatcherTest {

    @Test
    public void sheetNumbersAreEqual() throws Exception {
        assertThat(firstSheetOf("rowNumbersAreEqual.xls"), hasSameNumberOfRowAs(firstSheetOf("rowNumbersAreEqual.xls")));
    }

    @Test
    public void sheetNumbersAreNotEqual() throws Exception {
        assertThat(firstSheetOf("rowNumbersAreNotEqual.xls"), not(hasSameNumberOfRowAs(firstSheetOf("rowNumbersAreEqual.xls"))));
    }
}
