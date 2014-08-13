
import org.junit.Assert;
import org.junit.Test;

import java.io.File;

/**
 * Unit test for simple App.
 */
public class AppTest 

{

    /**
     * Rigourous Test :-)
     */
    @Test
    public void testAsposeWordExample()
    {
        final String WordOutPutFile = "output/AsposeMaven.doc";

        File wordfile = new File(WordOutPutFile);
        wordfile.delete();
        try {
            AsposeJavaExamples.runAsposeWord(WordOutPutFile);
        } catch (Exception e) {

            Assert.assertTrue(false);
        }

        Assert.assertTrue(wordfile.exists());
    }
}
