package bad.robot.excel;

import bad.robot.excel.valuetypes.Alignment;
import bad.robot.excel.valuetypes.Border;
import bad.robot.excel.valuetypes.DataFormat;
import bad.robot.excel.valuetypes.FontSize;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Workbook;

public class DefaultStyle implements Style {

    private final DataFormat format;
    private final Alignment alignment;
    private final FontSize fontSize;
    private final Border border;

    /** package protected. use {@link StyleBuilder} instead */
    DefaultStyle(Border border, DataFormat format, Alignment alignment, FontSize fontSize) {
        this.border = border;
        this.format = format;
        this.alignment = alignment;
        this.fontSize = fontSize;
    }

    @Override
    public void applyTo(org.apache.poi.ss.usermodel.Cell cell, Workbook workbook) {
        cell.setCellStyle(createStyle(workbook));
    }

    private CellStyle createStyle(Workbook workbook) {
        CellStyle style = workbook.createCellStyle();
        applyBorderTo(style);
        applyAlignmentTo(style);
        applyDataFormatTo(style, workbook);
        applyFontTo(style, workbook);
        return style;
    }

    private void applyBorderTo(CellStyle style) {
        if (border != null) {
            style.setBorderBottom(border.getBottom().value().getPoiStyle());
            style.setBorderLeft(border.getTop().value().getPoiStyle());
            style.setBorderRight(border.getRight().value().getPoiStyle());
            style.setBorderLeft(border.getLeft().value().getPoiStyle());
        }
    }

    private void applyAlignmentTo(CellStyle style) {
        if (alignment != null)
            style.setAlignment(alignment.value().getPoiStyle());
    }

    private void applyDataFormatTo(CellStyle style, Workbook workbook) {
        if (format != null)
            style.setDataFormat(workbook.createDataFormat().getFormat(format.value()));
    }

    private void applyFontTo(CellStyle style, Workbook workbook) {
        if (fontSize != null) {
            Font font = workbook.createFont();
            font.setFontHeightInPoints(fontSize.value());
            style.setFont(font);
        }
    }
}
