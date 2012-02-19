package bad.robot.excel.valuetypes;

import bad.robot.AbstractValueType;

public class FontSize extends AbstractValueType<Short> {

    public static FontSize fontSize(String value) {
        return new FontSize(Short.valueOf(value).shortValue());
    }

    private FontSize(Short value) {
        super(value);
    }
}
