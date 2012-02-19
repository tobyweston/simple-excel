package bad.robot.excel;

import org.apache.poi.ss.usermodel.CellStyle;

public enum BorderStyle {

    NONE(CellStyle.BORDER_NONE),
    THIN_SOLID(CellStyle.BORDER_THIN),
    MEDIUM_SOLID(CellStyle.BORDER_MEDIUM),
    THICK_SOLID(CellStyle.BORDER_THICK);

    private short poiStyle;

    BorderStyle(short poiStyle) {
        this.poiStyle = poiStyle;
    }

    public Short getPoiStyle() {
        return poiStyle;
    }
}
