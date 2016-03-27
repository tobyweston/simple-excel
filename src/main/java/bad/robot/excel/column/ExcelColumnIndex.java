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

package bad.robot.excel.column;

import org.junit.experimental.categories.Categories.ExcludeCategory;

/**
 * Currently supporting only 702 columns. sorry.
 */
public class ExcelColumnIndex {

	private int ordinal;

	private String value;

	private static ExcelColumnIndex[] values = {};

	private static final int MAX_COLUMNS = 16384;

	private ExcelColumnIndex(int ord, String val) {
		this.ordinal = ord;
		this.value = val;
	}

	public int ordinal() {
		return this.ordinal;
	}

	public String name() {
		return this.value;
	}

	private static ExcelColumnIndex[] getValues() {
		if (values.length == 0) {
			values = new ExcelColumnIndex[MAX_COLUMNS];
			int cnt = 0;
			for (String s: ColumnsGenerator.generateColumns(MAX_COLUMNS)) {
				if (cnt > (MAX_COLUMNS-1)) {
					break;
				}

				values[cnt] = new ExcelColumnIndex(cnt, s);

				cnt++;
			}
		}

		return values;
	}

	public static ExcelColumnIndex[] values() {
		return getValues();
	}

    public static ExcelColumnIndex from(int index) {
        return getValues()[index];
    }

    public static int valueOf(String key) {
    	for (ExcelColumnIndex s: getValues()) {
    		if (s.name().equals(key)) {
    			return s.ordinal();
    		}
    	}

    	throw new IllegalArgumentException("Column with key "+key+" does not exists");
    }

    public static ExcelColumnIndex getColumn(String key) {
    	for (ExcelColumnIndex s: getValues()) {
    		if (s.name().equals(key.toUpperCase())) {
    			return s;
    		}
    	}

    	throw new IllegalArgumentException("Column with key "+key+" does not exists");
    }

    private static class ColumnsGenerator {


    	private static char[] getLetters() {
    		char letters[] = new char[26];

    		int index = 0;
    		for (char i='A'; i<= 'Z'; i++) {
    			letters[index] = i;
    			index++;
    		}

    		return letters;
    	}

    	public static String getColumnByIndex(int index) {
    		return generateColumns(index+1)[index];
    	}

    	private static String[] generateColumns(int length) {

    		char characters[] = getLetters();
    		int charactersLength = characters.length,
    			callTimes = (int)(length / charactersLength);

    		String generated[] = new String[length+charactersLength];

    		if ((length % charactersLength) != 0)
    		    callTimes++;


    		int currentIndex = 0, skip = 0;
    		for (int i = 0; i<=callTimes; i++) {

    			if (generated[0] == null) {
    				for (char j: characters) {
    					generated[currentIndex] = j+"";
    					currentIndex ++;
    				}
    			} else {
    				int until = currentIndex,
    				    from,
    				    counter = 0;

                    firstFor:
    				for (from = skip; from < until; from++) {
    					for (char j: characters) {
    					    if (currentIndex > (generated.length-1)) {
    					        break firstFor;
    					    }

    						generated[currentIndex] = generated[from]+j;
    						currentIndex++;
    						counter++;
    					}
    				}

    				skip = currentIndex-counter;
    			}

    		}

    		return generated;
    	}
    }
}
