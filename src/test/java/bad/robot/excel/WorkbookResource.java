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

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Sheet;

import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;

import static bad.robot.excel.Assertions.assertNotNull;

public class WorkbookResource {

    public static HSSFWorkbook getWorkbook(String file) throws IOException {
        InputStream stream = PoiWorkbookMutatorTest.class.getResourceAsStream(file);
        if (stream == null)
            throw new FileNotFoundException(file);
        return new HSSFWorkbook(stream);
    }

    public static Sheet firstSheetOf(String file) throws IOException {
        return assertNotNull(getWorkbook(file).getSheetAt(0));
    }

    public static org.apache.poi.ss.usermodel.Row firstRowOf(String file) throws IOException {
        return assertNotNull(firstSheetOf(file).getRow(0));
    }

    public static org.apache.poi.ss.usermodel.Row secondRowOf(String file) throws IOException {
        return assertNotNull(firstSheetOf(file).getRow(1));
    }

    public static org.apache.poi.ss.usermodel.Row thirdRowOf(String file) throws IOException {
        return assertNotNull(firstSheetOf(file).getRow(2));
    }

}
