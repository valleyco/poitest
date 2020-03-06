/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package valleyco.msword;

import java.math.BigInteger;
import javax.xml.bind.annotation.adapters.HexBinaryAdapter;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFStyle;
import org.apache.poi.xwpf.usermodel.XWPFStyles;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTColor;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTDecimalNumber;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTFonts;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTHpsMeasure;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTOnOff;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTPPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTRPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTString;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTStyle;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STStyleType;

/**
 *
 * @author David Levy <levydav@gmail.com>
 */
public class StylePara {

    private String strStyleId;
    private int headingLevel = 1;
    private int pointSize = 10;
    private String hexColor = "0000";
    private String fontName = "";
    XWPFStyle style = null;

    public StylePara(String strStyleId) {
        this.strStyleId = strStyleId;
    }

    public XWPFStyle toStyle() {
        if (style == null) {
            style = createStyle();
        }
        return style;
    }

    public StylePara setHeadingLevel(int headingLevel) {
        this.headingLevel = headingLevel;
        return this;
    }

    public StylePara setPointSize(int pointSize) {
        this.pointSize = pointSize;
        return this;
    }

    public StylePara setColor(String color) {
        this.hexColor = color;
        return this;
    }

    public StylePara setFont(String name) {
        this.fontName = name;
        return this;
    }

    private XWPFStyle createStyle() {

        CTStyle ctStyle = CTStyle.Factory.newInstance();
        ctStyle.setStyleId(strStyleId);

        ctStyle.setName(CTFactory.getString(strStyleId));

        // lower number > style is more prominent in the formats bar
        ctStyle.setUiPriority(CTFactory.getDecimalNumber(headingLevel));

        CTOnOff onoffnull = CTOnOff.Factory.newInstance();
        ctStyle.setUnhideWhenUsed(onoffnull);

        // style shows up in the formats bar
        ctStyle.setQFormat(onoffnull);

        // style defines a heading of the given level
        CTPPr ppr = CTPPr.Factory.newInstance();

        ppr.setOutlineLvl(CTFactory.getDecimalNumber(headingLevel));
        ctStyle.setPPr(ppr);

        XWPFStyle pfStyle = new XWPFStyle(ctStyle);

        CTHpsMeasure size = CTHpsMeasure.Factory.newInstance();
        size.setVal(new BigInteger(String.valueOf(pointSize)));
        CTHpsMeasure size2 = CTHpsMeasure.Factory.newInstance();
        size2.setVal(new BigInteger("24"));

        CTFonts fonts = CTFonts.Factory.newInstance();
        fonts.setAscii(fontName);

        CTRPr rpr = CTRPr.Factory.newInstance();
        rpr.setRFonts(fonts);
        rpr.setSz(size);
        rpr.setSzCs(size2);

        CTColor color = CTColor.Factory.newInstance();
        color.setVal(hexToBytes(hexColor));
        rpr.setColor(color);
        pfStyle.getCTStyle().setRPr(rpr);
        // is a null op if already defined

        pfStyle.setType(STStyleType.PARAGRAPH);
        return pfStyle;

    }

    public static byte[] hexToBytes(String hexString) {
        HexBinaryAdapter adapter = new HexBinaryAdapter();
        byte[] bytes = adapter.unmarshal(hexString);
        return bytes;
    }
}
