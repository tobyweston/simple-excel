package bad.robot.excel;

import bad.robot.excel.valuetypes.Alignment;
import bad.robot.excel.valuetypes.Border;
import bad.robot.excel.valuetypes.DataFormat;
import bad.robot.excel.valuetypes.FontSize;

public class StyleBuilder {
    
    private DataFormat format;
    private Alignment alignment;
    private FontSize fontSize;
    private Border border;

    private StyleBuilder() {
    }

    public static StyleBuilder aStyle() {
        return new StyleBuilder();
    }

    public StyleBuilder with(DataFormat format) {
        this.format = format;
        return this;
    }

    public StyleBuilder with(Alignment alignment) {
        this.alignment = alignment;
        return this;
    }

    public StyleBuilder with(FontSize fontSize) {
        this.fontSize = fontSize;
        return this;
    }

    public StyleBuilder with(Border border) {
        this.border = border;
        return this;
    }

    public DefaultStyle build() {
        return new DefaultStyle(border, format, alignment, fontSize);
    }
}