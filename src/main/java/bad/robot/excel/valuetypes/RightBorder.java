package bad.robot.excel.valuetypes;

import bad.robot.AbstractValueType;
import bad.robot.excel.BorderStyle;

public class RightBorder extends AbstractValueType<BorderStyle> {

    private RightBorder(BorderStyle value) {
        super(value);
    }

    public static RightBorder right(BorderStyle value) {
        return new RightBorder(value);
    }
}
