package org.docx4j.toc;

import java.io.File;
import junit.framework.Assert;

import org.docx4j.TraversalUtil;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.docx4j.wml.Body;
import org.docx4j.wml.Document;
import org.docx4j.wml.SdtBlock;
import org.junit.Test;
import org.docx4j.wml.P;
import org.docx4j.wml.R;
import javax.xml.bind.JAXBElement;

public class TocGenerateTest {

    public static final String TOC_STYLE_MASK = "TOC%s";
    final static String PATH = "C:\\Users\\a940\\Desktop\\"; 
    static String inputLectura = PATH+"contabla.docx";	
	
    @Test
    public void testPropio() throws Exception{
        WordprocessingMLPackage WMLP = WordprocessingMLPackage.load(new File(inputLectura));
        SdtBlock sdt = getTocSDT(WMLP);
        int size = sdt.getSdtContent().getContent().size();
        System.out.println("testPropio size: " + size);
        
        for(int i=0;i<size;i++){
            P pe = (P) sdt.getSdtContent().getContent().get(i);
            int subsize = pe.getContent().size();
            System.out.println("testPropio elem "+i+": " + pe.getContent()); 
            System.out.println("testPropio elem "+i+": size " + pe.getContent().size()); 
            
            for(int j=0;j<subsize;j++){
                if(pe.getContent().get(j).getClass() == R.class){
                    R erre = (R) pe.getContent().get(j);
                    System.out.println("subelem R "+i+"."+j+": " + erre.getContent());
                }
                else if(pe.getContent().get(j).getClass() == JAXBElement.class){
                    JAXBElement jaxb = (JAXBElement) pe.getContent().get(j);
                    System.out.println("subelem JaxB "+i+"."+j+": " + jaxb.getName() + " - " + jaxb.getValue());
                }
                else
                    System.out.println("subelem else "+i+"."+j+": " + pe.getContent().get(j));
            }
        }
    }
    
//    @Test
    public void testGeneral() throws TocException, Exception {

//        WordprocessingMLPackage WMLP = createPkg();
//        
//        TocGenerator tocGenerator = new TocGenerator(WMLP);
//        tocGenerator.generateToc( 0, "TOC \\o \"1-3\" \\h \\z \\u ", true);
//        SdtBlock sdt = getTocSDT(WMLP);
//
//        int size = sdt.getSdtContent().getContent().size();
//        System.out.println("testGeneral size: " + size);
//        for(int i=0;i<size;i++){
//            System.out.println("testGeneral elem "+i+": " + sdt.getSdtContent().getContent().get(i).toString()); 
//        }
//        System.out.println("testGeneral xml: " + wordMLPackage.getMainDocumentPart().getXML());
//
//        /* Title p + instruction p +  3 entries + end p */
//        Assert.assertEquals(6, sdt.getSdtContent().getContent().size());
    }

//    @Test
    public void testHeading() throws TocException, Exception {

//        WordprocessingMLPackage wordMLPackage = createPkg();
//
//        TocGenerator tocGenerator = new TocGenerator(wordMLPackage);
//        tocGenerator.generateToc( 0, "TOC \\o \"1-3\"", true);
//
//
//        SdtBlock sdt = getTocSDT(wordMLPackage);
//        
//        System.out.println("testHeading size: " + sdt.getSdtContent().getContent().size());        
//        System.out.println(wordMLPackage.getMainDocumentPart().getXML());
//                
//        /* Title p + instruction p +  3 entries + end p */
//        Assert.assertEquals(6, sdt.getSdtContent().getContent().size());
    }

//    @Test
    public void testHyperlink() throws TocException, Exception {

//          WordprocessingMLPackage wordMLPackage = createPkg();
//		
//          TocGenerator.generateToc(wordMLPackage, 0, "TOC \\h", true);        
//        
//          SdtBlock sdt = getTocSDT(wordMLPackage);
//                        
//          /* Title p + instruction p +  3 entries + end p */
//          Assert.assertEquals(6, sdt.getSdtContent().getContent().size());
    }

//    @Test
    public void testOutlineLevel() throws TocException, Exception {

//            WordprocessingMLPackage wordMLPackage = createPkg();
//		
//            TocGenerator tocGenerator = new TocGenerator(wordMLPackage);
//            tocGenerator.generateToc( 0, "TOC \\u", true);
//        
//            SdtBlock sdt = getTocSDT(wordMLPackage);
//        
//            System.out.println(sdt.getSdtContent().getContent().size());        
//            System.out.println(wordMLPackage.getMainDocumentPart().getXML());
//            System.out.println(wordMLPackage.getMainDocumentPart().getStyleDefinitionsPart().getXML());
//                
//            /*Title p + instruction p +  3 entries + end p*/
//            Assert.assertEquals(6, sdt.getSdtContent().getContent().size());
    }


//    @Test
    public void testHeadingTrumpsOutline() throws TocException, Exception {

//          WordprocessingMLPackage wordMLPackage = createPkg();
//	
//          TocGenerator tocGenerator = new TocGenerator(wordMLPackage);
//          tocGenerator.generateToc( 0, "TOC \\o \"1-2\" \\u", true);
//
//          SdtBlock sdt = getTocSDT(wordMLPackage);
//
//          System.out.println(sdt.getSdtContent().getContent().size());        
//          System.out.println(wordMLPackage.getMainDocumentPart().getXML());
//
//            /*Title p + instruction p +  2 entries + end p*/
//          Assert.assertEquals(5, sdt.getSdtContent().getContent().size());
    }

//    @Test
    public void testToCHeadingNull() throws TocException, Exception {

//          Toc.setTocHeadingText(null); // Word is ok with this
//		
//          WordprocessingMLPackage wordMLPackage = createPkg();
//		
//          TocGenerator tocGenerator = new TocGenerator(wordMLPackage);
//          tocGenerator.generateToc( 0, " TOC \\o \"1-3\" \\h \\z \\u ", true);
//        
//          SdtBlock sdt = getTocSDT(wordMLPackage);
//        
//          System.out.println(sdt.getSdtContent().getContent().size());        
//          System.out.println(wordMLPackage.getMainDocumentPart().getXML());
//        
//          Docx4J.save(wordMLPackage, new File("testToCHeadingNull.docx"));
    }

//    @Test
    public void testToCHeadingEmpty() throws TocException, Exception {

//          Toc.setTocHeadingText(""); // Word is ok with this
//		
//          WordprocessingMLPackage wordMLPackage = createPkg();
//		
//          TocGenerator tocGenerator = new TocGenerator(wordMLPackage);
//          tocGenerator.generateToc( 0, " TOC \\o \"1-3\" \\h \\z \\u ", true);
//        
//          SdtBlock sdt = getTocSDT(wordMLPackage);
//        
//          System.out.println(sdt.getSdtContent().getContent().size());        
//          System.out.println(wordMLPackage.getMainDocumentPart().getXML());
//        
//          Docx4J.save(wordMLPackage, new File("testToCHeadingEmpty.docx"));
    }

//    @Test
    public void testToCHeadingSet() throws TocException, Exception {

//          Toc.setTocHeadingText("Alpha");
//		
//          WordprocessingMLPackage wordMLPackage = createPkg();
//		
//          TocGenerator tocGenerator = new TocGenerator(wordMLPackage);
//          tocGenerator.generateToc( 0, " TOC \\o \"1-3\" \\h \\z \\u ", true);
//        
//          SdtBlock sdt = getTocSDT(wordMLPackage);
//        
//          System.out.println(sdt.getSdtContent().getContent().size());        
//          System.out.println(wordMLPackage.getMainDocumentPart().getXML());
//        
//          Docx4J.save(wordMLPackage, new File("testToCHeadingSet.docx"));
    }
	
    private SdtBlock getTocSDT(WordprocessingMLPackage wordMLPackage) {

        MainDocumentPart documentPart = wordMLPackage.getMainDocumentPart();
        Document wmlDocumentEl = (Document)documentPart.getJaxbElement();
        Body body =  wmlDocumentEl.getBody();

    	TocFinder finder = new TocFinder();
		new TraversalUtil(body.getContent(), finder);
		
		return finder.tocSDT;		
	}
    
    private WordprocessingMLPackage createPkg() throws Exception{
    	
        WordprocessingMLPackage wordMLPackage = WordprocessingMLPackage.createPackage();
        MainDocumentPart documentPart = wordMLPackage.getMainDocumentPart();
        
        for(int i = 1; i < 10; i++){
            documentPart.getPropertyResolver().activateStyle(String.format(TOC_STYLE_MASK, i));
        }
        
        documentPart.addStyledParagraphOfText("Heading1", "Hello 1");
        fillPageWithContent(documentPart, "Hello 1");
        documentPart.addStyledParagraphOfText("Heading1", ""); // Word omits empty entries from ToC
        fillPageWithContent(documentPart, "Hello 1");
        documentPart.addStyledParagraphOfText("Heading2", "Hello 2");
        fillPageWithContent(documentPart, "Hello 2");
        documentPart.addStyledParagraphOfText("Heading3", "Hello 3");
        fillPageWithContent(documentPart, "Hello 3");
        
        return wordMLPackage;
    }

    private static void fillPageWithContent(MainDocumentPart documentPart, String content){
        for(int i = 0; i < 2; i++){
            documentPart.addStyledParagraphOfText("Normal", content + " paragraph " + i);
        }
    }
}
