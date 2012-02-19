package bad.robot.excel;

import org.apache.poi.ss.usermodel.CellStyle;

public enum AlignmentStyle {

    LEFT(CellStyle.ALIGN_LEFT),
    CENTRE(CellStyle.ALIGN_CENTER),
    RIGHT(CellStyle.ALIGN_RIGHT),
    JUSTIFY(CellStyle.ALIGN_JUSTIFY);

    private short poiStyle;

    AlignmentStyle(short poiStyle) {
        this.poiStyle = poiStyle;
    }

    public Short getPoiStyle() {
        return poiStyle;
    }
}
