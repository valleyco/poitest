/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package valleyco.poitest;

/**
 *
 * @author david
 */
import org.apache.poi.util.Units;
import org.apache.poi.xwpf.usermodel.*;

import java.io.FileOutputStream;
import java.io.IOException;
import java.net.URISyntaxException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.stream.Collectors;
import java.util.stream.Stream;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTP;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTPPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTbl;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTblPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STOnOff;
import valleyco.msword.StylePara;

public class WordDocument {

    public static String logo = "logo-leaf.png";
    public static String logo_template = "logo.docx";
    public static String paragraph1 = "poi-word-para1.txt";
    public static String paragraph2 = "poi-word-para2.txt";
    public static String paragraph3 = "poi-word-para3.txt";
    public static String output = "rest-with-spring.docx";

    public void handleSimpleDoc() throws Exception {
        //XWPFDocument document = new XWPFDocument();
        XWPFDocument document = new XWPFDocument( Files.newInputStream(Paths.get(ClassLoader.getSystemResource(logo_template).toURI())));
        XWPFParagraph title = document.createParagraph();
        title.setAlignment(ParagraphAlignment.CENTER);
        XWPFRun titleRun = title.createRun();
        titleRun.setText("Build Your REST API with Spring");
        titleRun.setColor("009933");
        titleRun.setBold(true);
        titleRun.setFontFamily("Courier");
        titleRun.setFontSize(20);

        XWPFParagraph subTitle = document.createParagraph();
        subTitle.setAlignment(ParagraphAlignment.CENTER);
        XWPFRun subTitleRun = subTitle.createRun();
        subTitleRun.setText("from HTTP fundamentals to API Mastery");
        subTitleRun.setColor("00CC44");
        subTitleRun.setFontFamily("Courier");
        subTitleRun.setFontSize(16);
        subTitleRun.setTextPosition(20);
        subTitleRun.setUnderline(UnderlinePatterns.DOT_DOT_DASH);

        XWPFParagraph imageParagraph = document.createParagraph();
        imageParagraph.setAlignment(ParagraphAlignment.CENTER);
        XWPFRun imageRun = imageParagraph.createRun();
        imageRun.setTextPosition(20);
        Path imagePath = Paths.get(ClassLoader.getSystemResource(logo).toURI());
        imageRun.addPicture(
                Files.newInputStream(imagePath),
                XWPFDocument.PICTURE_TYPE_PNG,
                imagePath.getFileName().toString(),
                Units.toEMU(50),
                Units.toEMU(50)
        );

        XWPFParagraph sectionTitle = document.createParagraph();
        XWPFRun sectionTRun = sectionTitle.createRun();
        sectionTRun.setText("What makes a good API?");
        sectionTRun.setColor("00CC44");
        sectionTRun.setBold(true);
        sectionTRun.setFontFamily("Courier");

        XWPFParagraph para1 = document.createParagraph();
        para1.setAlignment(ParagraphAlignment.BOTH);
        String string1 = convertTextFileToString(paragraph1);
        XWPFRun para1Run = para1.createRun();
        para1Run.setText(string1);

        XWPFParagraph para2 = document.createParagraph();
        para2.setAlignment(ParagraphAlignment.RIGHT);
        String string2 = convertTextFileToString(paragraph2);
        XWPFRun para2Run = para2.createRun();
        para2Run.setText(string2);
        para2Run.setItalic(true);

        XWPFParagraph para3 = document.createParagraph();
        para3.setAlignment(ParagraphAlignment.LEFT);
        String string3 = convertTextFileToString(paragraph3);
        XWPFRun para3Run = para3.createRun();
        para3Run.setText(string3);

        var style = (new StylePara("test"))
                .setColor("0000ff")
                .setPointSize(25)
                .setFont("David");
        document.createStyles().addStyle(style.toStyle());

        XWPFParagraph para4 = document.createParagraph();
        CTP ctPara4 = para4.getCTP();
        CTPPr ctppr4;
        if ((ctppr4 = ctPara4.getPPr()) == null) {
            ctppr4 = ctPara4.addNewPPr();
        }
        ctppr4.addNewBidi().setVal(STOnOff.ON);

        para4.setStyle("test");
        para4.setAlignment(ParagraphAlignment.LEFT);
        XWPFRun para4Run = para4.createRun();
        para4Run.setText(string3);

//create table
        XWPFTable table = document.createTable();
        XWPFTableCell cell;

        CTTbl ctt = table.getCTTbl();
        CTTblPr ctppr;

        if ((ctppr = ctt.getTblPr()) == null) {
            ctppr = ctt.addNewTblPr();
        }
        ctppr.addNewBidiVisual().setVal(STOnOff.ON);

        //create first row
        XWPFTableRow tableRowOne = table.getRow(0);
        tableRowOne.getCell(0).setText("col one, row one");
        tableRowOne.addNewTableCell().setText("col two, row one");
        tableRowOne.addNewTableCell().setText("col three, row one");

        //create second row
        XWPFTableRow tableRowTwo = table.createRow();
        tableRowTwo.getCell(0).setText("col one, row two");
        tableRowTwo.getCell(1).setText("col two, row two");
        tableRowTwo.getCell(2).setText("col three, row two");

        //create third row
        XWPFTableRow tableRowThree = table.createRow();
        cell = tableRowThree.getCell(0);

        cell.setText("col one, row three");
        tableRowThree.getCell(1).setText("col two, row three");
        tableRowThree.getCell(2).setText("col three, row three");

        FileOutputStream out = new FileOutputStream(output);
        document.write(out);
        out.close();
        document.close();
    }

    public static String convertTextFileToString(String fileName) {
        try (Stream<String> stream = Files.lines(Paths.get(ClassLoader.getSystemResource(fileName).toURI()))) {
            return stream.collect(Collectors.joining(" "));
        } catch (IOException | URISyntaxException e) {
            e.printStackTrace();
        }
        return null;
    }

}
