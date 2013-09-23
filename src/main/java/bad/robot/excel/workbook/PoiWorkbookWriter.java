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

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;

import static java.io.File.createTempFile;

public class PoiWorkbookWriter implements Writable {

    private final Workbook workbook;

    public PoiWorkbookWriter(Workbook workbook) {
        this.workbook = workbook;
    }

    @Override
    public void writeTo(OutputStream stream) {
        try {
            workbook.write(stream);
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }

    @Override
    public void saveAs(File file) {
        try {
            FileOutputStream stream = new FileOutputStream(file);
            workbook.write(stream);
            stream.close();
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }

    @Override
    public void saveAsTemporaryFile(String filenameExcludingExtension) {
        String extension = workbook instanceof HSSFWorkbook ? ".xls" : ".xlsx";
        try {
            File file = createTempFile(filenameExcludingExtension, extension);
            saveAs(file);
            System.out.printf("saved %s: %s%n", extension, file.getAbsolutePath());
        } catch (IOException e) {
            throw new RuntimeException(String.format("failed to save %s: %s.%s", extension, filenameExcludingExtension, extension), e);
        }

    }
}
