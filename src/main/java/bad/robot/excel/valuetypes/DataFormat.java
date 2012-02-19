package bad.robot.excel.valuetypes;

import bad.robot.AbstractValueType;

public class DataFormat extends AbstractValueType<String> {

    public static DataFormat dataFormat(String value) {
        return new DataFormat(value);
    }

    private DataFormat(String value) {
        super(value);
    }

}
