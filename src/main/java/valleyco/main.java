package valleyco;

import java.io.File;
import java.nio.file.Paths;
import java.util.logging.Level;
import java.util.logging.Logger;
import valleyco.poitest.SimpleTable;
import valleyco.poitest.WordDocument;
import static valleyco.poitest.WordDocument.logo;


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
    
    public static void main(String args[]) throws Exception {
        poiDoc();
        SimpleTable.demo(args);
    }
}
