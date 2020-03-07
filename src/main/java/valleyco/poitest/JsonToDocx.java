/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package valleyco.poitest;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Iterator;
import org.apache.poi.xwpf.usermodel.ParagraphAlignment;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.json.JSONArray;
import org.json.JSONObject;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTSimpleField;

/**
 *
 * @author David Levy <levydav@gmail.com>
 */
public class JsonToDocx {

    final private XWPFDocument document;
    final JSONObject jsonDoc;

    public JsonToDocx(JSONObject jsonDoc) {
        this.jsonDoc = jsonDoc;
        document = new XWPFDocument();
    }

    public JsonToDocx(JSONObject jsonDoc, XWPFDocument document) {
        this.jsonDoc = jsonDoc;
        this.document = document;
    }

    public void convert() {
        JSONArray sections = (JSONArray) jsonDoc.get("sections");
        for (Object o : sections) {
            JSONObject section = (JSONObject) o;
            String type = section.getString("type");
            switch (type) {
                case "para":
                    addPara(section);
                    break;
                case "table":
                    addTable(section);
                    break;
            }
        }
    }

    private void addPara(JSONObject para) {
        XWPFParagraph xpara = document.createParagraph();
        JSONArray runs = new JSONArray();
        Iterator<String> keys = para.keys();

        while (keys.hasNext()) {
            String key = keys.next();
            switch (key) {
                case "type":
                    break;
                case "style":
                    xpara.setStyle(para.getString(key));
                    break;
                case "Alignment":
                    xpara.setAlignment(ParagraphAlignment.valueOf(para.getString(key).toUpperCase()));
                    break;
                case "PageBreak":
                    xpara.setPageBreak(para.getBoolean(key));
                    break;
                case "SpacingAfter":
                    xpara.setSpacingAfter(para.getInt(key));
                    break;
                case "SpacingBefore":
                    xpara.setSpacingBefore(para.getInt(key));
                    break;
                case "Runs":
                    runs = para.getJSONArray(key);
                    break;
            }
        }
        for (Object o : runs) {
            addParaRun(xpara, (JSONObject) o);
        }
    }

    private void addParaRun(XWPFParagraph p, JSONObject run) {
        switch (run.getString("Type")) {
            case "Text":
                addParaTextRun(p, run);
                break;
            case "SimpleField":
                addField(p, run);
                break;
        }
    }

    private void addParaTextRun(XWPFParagraph p, JSONObject run) {

        XWPFRun paraRun = p.createRun();
        Iterator<String> runKeys = run.keys();
        while (runKeys.hasNext()) {
            String key = runKeys.next();
            switch (key) {
                case "Text":
                    paraRun.setText(run.getString(key));
                    break;
                case "Style":
                    paraRun.setStyle(run.getString(key));
                    break;
                case "FontSize":
                    paraRun.setFontSize(run.getInt(key));
                    break;
                case "FontColor":
                    paraRun.setColor(run.getString(key));
                case "FontFamily":
                    paraRun.setFontFamily(run.getString(key));
                    break;
                case "TextPosition":
                    paraRun.setTextPosition(run.getInt(key));
                    break;
            }

        }

    }

    private static void addField(XWPFParagraph paragraph, JSONObject run) {
        String fieldName = run.getString("Name");
        String text = run.has("Text") ? run.getString("Text") : "<<" + fieldName + ">>";
        CTSimpleField ctSimpleField = paragraph.getCTP().addNewFldSimple();
        ctSimpleField.setInstr(fieldName + " \\* MERGEFORMAT ");
        ctSimpleField.addNewR().addNewT().setStringValue(text);
    }

    private void addTable(JSONObject table) {

    }

    public void save(String filename) throws FileNotFoundException, IOException {
        FileOutputStream out = new FileOutputStream(filename);
        document.write(out);
        out.close();
        document.close();
    }
}
