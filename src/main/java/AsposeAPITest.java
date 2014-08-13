import com.aspose.words.*;

import java.util.Iterator;

/**
 * Created by Adeel Ilyas on 6/27/2014.
 */

/*

 Aspose.Words Example
 *
 */
public class AsposeAPITest {

    public static void main( String args[]) throws Exception {
        final String WordOutPutFile = "output/AsposeMaven.doc";
        System.out.println("Welcome to Aspose For Maven!");

        try {
            AsposeJavaExamples.runAsposeWord(WordOutPutFile);
        } catch (Exception e) {
            throw new Exception("Aspose: Unable to export to ms word format.. some error occured",e);

        }

    }

}
