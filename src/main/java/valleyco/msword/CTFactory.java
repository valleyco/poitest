/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package valleyco.msword;

import java.math.BigInteger;
import org.apache.xmlbeans.XmlException;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTDecimalNumber;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTString;

/**
 *
 * @author david
 */
public class CTFactory {

    public static CTString getString(String str) {
        CTString result = CTString.Factory.newInstance();
        result.setVal(str);
        return result;
    }

    public static CTDecimalNumber getDecimalNumber(String string) throws XmlException {
        CTDecimalNumber result = CTDecimalNumber.Factory.parse(string);
        return result;
    }

    public static CTDecimalNumber getDecimalNumber(long n) {
        CTDecimalNumber result = CTDecimalNumber.Factory.newInstance();
        result.setVal(BigInteger.valueOf(n));
        return result;
    }
}
