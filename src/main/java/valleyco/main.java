package valleyco;

import java.io.IOException;
import java.net.URISyntaxException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.logging.Level;
import java.util.logging.Logger;
import org.json.JSONObject;
import org.json.JSONTokener;
import valleyco.poitest.SimpleTable;
import valleyco.poitest.WordDocument;


/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
/**
 *
 * @author david
 */
public class main {

    private static void poiDoc() {
        var test = new WordDocument();
        try {
            test.handleSimpleDoc();
        } catch (Exception ex) {
            Logger.getLogger(WordDocument.class.getName()).log(Level.SEVERE, null, ex);
        }
    }
    private static JSONObject getJson(String filename) throws URISyntaxException, IOException {
        Path filePath = Paths.get(ClassLoader.getSystemResource(filename).toURI());
        JSONTokener tok = new JSONTokener(Files.newInputStream(filePath));
        return new JSONObject(tok);
    }
    public static void main(String args[]) throws Exception {
        poiDoc();
        SimpleTable.demo(args);
    }
}
