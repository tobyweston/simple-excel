package bad.robot.excel.valuetypes;

import bad.robot.AbstractValueType;
import bad.robot.excel.BorderStyle;

public class BottomBorder extends AbstractValueType<BorderStyle> {

    private BottomBorder(BorderStyle value) {
        super(value);
    }

    public static BottomBorder bottom(BorderStyle value) {
        return new BottomBorder(value);
    }
}
