/*
 * Copyright (c) 2012-2013, bad robot (london) ltd.
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *     http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */

package bad.robot.excel.workbook;

import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;

public class PoiWorkbookReader implements Readable<Workbook> {

    @Override
    public Workbook read(File file) throws IOException {
        InputStream stream = new FileInputStream(file);
        return read(stream);
    }

    @Override
    public Workbook read(InputStream stream) throws IOException {
        if (stream == null)
            throw new IllegalArgumentException("stream was null, could not load the workbook");
        try {
            return WorkbookFactory.create(stream);
        } finally {
            stream.close();
        }
    }

}
