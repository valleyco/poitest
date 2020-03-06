/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package valleyco.poitest;

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.json.JSONArray;
import org.json.JSONObject;

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

    }

    private void addTable(JSONObject table) {

    }

}
