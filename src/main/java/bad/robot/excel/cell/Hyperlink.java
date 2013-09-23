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

package bad.robot.excel.cell;

import bad.robot.excel.AbstractValueType;

import java.net.MalformedURLException;
import java.net.URL;

import static java.lang.String.format;

public class Hyperlink extends AbstractValueType<URL> {

    private final String text;

    public static Hyperlink hyperlink(String text, URL link) {
        return new Hyperlink(text, link);
    }

    public static Hyperlink hyperlink(String text, String url) {
        try {
            return new Hyperlink(text, new URL(url));
        } catch (MalformedURLException e) {
            throw new IllegalArgumentException(format("%s is not a valid URL", url), e);
        }
    }

    private Hyperlink(String text, URL url) {
        super(url);
        this.text = text;
    }

    public String text() {
        return text;
    }
}
