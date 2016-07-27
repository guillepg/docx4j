package org.docx4j.toc;

import java.io.File;
import java.util.logging.Level;
import java.util.logging.Logger;
import javax.xml.bind.JAXBElement;
import junit.framework.Assert;
import org.docx4j.Docx4J;
import org.junit.Test;

import org.docx4j.TraversalUtil;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.docx4j.wml.Body;
import org.docx4j.wml.Document;
import org.docx4j.wml.SdtBlock;
import org.docx4j.wml.P;
import org.docx4j.wml.R;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.wml.P.Hyperlink;
import org.docx4j.wml.Text;

public class TocGenerateTest {

    static final String TOC_STYLE_MASK = "TOC%s";
    final static String PATH = "C:/Users/a940/Desktop/CSCAE/"; 
    static String inputLectura = PATH + "contabla.docx";	
    static String inputfilepath = PATH + "antes.docx";
    static String outputfilepath = PATH + "despues.docx";
	
//    @Test
    public void testPropioLectura(){
        try {
            WordprocessingMLPackage WMLP = WordprocessingMLPackage.load(new File(inputLectura));
//            SdtBlock sdt = getTocSDT(WMLP);
//            leerBloqueSdt(sdt);
            System.out.println("Fichero en formato XML: " + WMLP.getMainDocumentPart().getXML());
        } catch (Docx4JException ex) {
            Logger.getLogger(TocGenerateTest.class.getName()).log(Level.INFO, null, "ERROR EN LECTURA");
            Logger.getLogger(TocGenerateTest.class.getName()).log(Level.SEVERE, null, ex);
        }
    }
    
//    @Test
    public void testPropioEscritura(){
        try {
            WordprocessingMLPackage WMLP = WordprocessingMLPackage.load(new File(inputfilepath));
            TocGenerator tocGen = new TocGenerator(WMLP);
            tocGen.generateToc(0,"TOC \\o \"1-3\" \\h \\z \\u ",false);
            Logger.getLogger(TocGenerateTest.class.getName()).log(Level.INFO, "Escritura OK", "");
        } catch (Docx4JException ex) {
            Logger.getLogger(TocGenerateTest.class.getName()).log(Level.INFO, "ERROR EN ESCRITURA", "");
            Logger.getLogger(TocGenerateTest.class.getName()).log(Level.SEVERE, null, ex);
        }
    }
    
    @Test
    public void testPropioActualizacion(){
        try {
            WordprocessingMLPackage WMLP = WordprocessingMLPackage.load(new File(inputLectura));
            TocGenerator tocGen = new TocGenerator(WMLP);
            SdtBlock tabla = tocGen.updateToc(false);
            
            Logger.getLogger(TocGenerateTest.class.getName()).log(Level.INFO, "Actualizacion OK", "");
//            leerBloqueSdt(tabla);

            P pe = (P) tabla.getSdtContent().getContent().get(2);
            JAXBElement jaxb = (JAXBElement) pe.getContent().get(0);
            Hyperlink link = (Hyperlink) jaxb.getValue();
            R erre = (R) link.getContent().get(0);
            Text texto = (Text) erre.getContent().get(0);
            texto.setValue(texto.getValue()+" NUEVO");
            System.out.println("Texto: "+ texto.getValue());
            
            tocGen.generateToc(tabla, "TOC \\o \"1-3\" \\h \\z \\u ", false);
        } catch (Exception ex) {
            Logger.getLogger(TocGenerateTest.class.getName()).log(Level.INFO, "ERROR EN ACTUALIZACION", "");
            Logger.getLogger(TocGenerateTest.class.getName()).log(Level.SEVERE, null, ex);
        }
    }
    
    private void leerBloqueSdt(SdtBlock sdt){
        int size = sdt.getSdtContent().getContent().size();
        System.out.println("leerBloqueSdt size: " + size); 

        for(int i=0;i<size;i++){
            P pe = (P) sdt.getSdtContent().getContent().get(i);
            int subsize = pe.getContent().size();
            System.out.println("testPropioLectura elem "+i+": " + pe.getContent());
            System.out.println("testPropioLectura elem "+i+": size " + pe.getContent().size());

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

        WordprocessingMLPackage WMLP = createPkg();
        
        TocGenerator tocGenerator = new TocGenerator(WMLP);
        tocGenerator.generateToc( 0, "TOC \\o \"1-3\" \\h \\z \\u ", true);
        SdtBlock sdt = getTocSDT(WMLP);

        int size = sdt.getSdtContent().getContent().size();
        System.out.println("testGeneral size: " + size);
        for(int i=0;i<size;i++){
            System.out.println("testGeneral elem "+i+": " + sdt.getSdtContent().getContent().get(i).toString()); 
        }
        System.out.println("testGeneral xml: " + WMLP.getMainDocumentPart().getXML());

        /* Title p + instruction p +  3 entries + end p */
        Assert.assertEquals(6, sdt.getSdtContent().getContent().size());
    }

//    @Test
    public void testHeading() throws TocException, Exception {

        WordprocessingMLPackage wordMLPackage = createPkg();

        TocGenerator tocGenerator = new TocGenerator(wordMLPackage);
        tocGenerator.generateToc( 0, "TOC \\o \"1-3\"", true);

        SdtBlock sdt = getTocSDT(wordMLPackage);
        
        System.out.println("testHeading size: " + sdt.getSdtContent().getContent().size());        
        System.out.println(wordMLPackage.getMainDocumentPart().getXML());
                
        /* Title p + instruction p +  3 entries + end p */
        Assert.assertEquals(6, sdt.getSdtContent().getContent().size());
    }

//    @Test
    public void testHyperlink() throws TocException, Exception {

        WordprocessingMLPackage wordMLPackage = createPkg();

        TocGenerator.generateToc(wordMLPackage, 0, "TOC \\h", true);        

        SdtBlock sdt = getTocSDT(wordMLPackage);

        /* Title p + instruction p +  3 entries + end p */
        Assert.assertEquals(6, sdt.getSdtContent().getContent().size());
    }

//    @Test
    public void testOutlineLevel() throws TocException, Exception {

        WordprocessingMLPackage wordMLPackage = createPkg();

        TocGenerator tocGenerator = new TocGenerator(wordMLPackage);
        tocGenerator.generateToc( 0, "TOC \\u", true);

        SdtBlock sdt = getTocSDT(wordMLPackage);

        System.out.println(sdt.getSdtContent().getContent().size());        
        System.out.println(wordMLPackage.getMainDocumentPart().getXML());
        System.out.println(wordMLPackage.getMainDocumentPart().getStyleDefinitionsPart().getXML());

        /*Title p + instruction p +  3 entries + end p*/
        Assert.assertEquals(6, sdt.getSdtContent().getContent().size());
    }


//    @Test
    public void testHeadingTrumpsOutline() throws TocException, Exception {

        WordprocessingMLPackage wordMLPackage = createPkg();

        TocGenerator tocGenerator = new TocGenerator(wordMLPackage);
        tocGenerator.generateToc( 0, "TOC \\o \"1-2\" \\u", true);

        SdtBlock sdt = getTocSDT(wordMLPackage);

        System.out.println(sdt.getSdtContent().getContent().size());        
        System.out.println(wordMLPackage.getMainDocumentPart().getXML());

          /*Title p + instruction p +  2 entries + end p*/
        Assert.assertEquals(5, sdt.getSdtContent().getContent().size());
    }

//    @Test
    public void testToCHeadingNull() throws TocException, Exception {

        Toc.setTocHeadingText(null); // Word is ok with this

        WordprocessingMLPackage wordMLPackage = createPkg();

        TocGenerator tocGenerator = new TocGenerator(wordMLPackage);
        tocGenerator.generateToc( 0, " TOC \\o \"1-3\" \\h \\z \\u ", true);

        SdtBlock sdt = getTocSDT(wordMLPackage);

        System.out.println(sdt.getSdtContent().getContent().size());        
        System.out.println(wordMLPackage.getMainDocumentPart().getXML());

        Docx4J.save(wordMLPackage, new File("testToCHeadingNull.docx"));
    }

//    @Test
    public void testToCHeadingEmpty() throws TocException, Exception {

        Toc.setTocHeadingText(""); // Word is ok with this

        WordprocessingMLPackage wordMLPackage = createPkg();

        TocGenerator tocGenerator = new TocGenerator(wordMLPackage);
        tocGenerator.generateToc( 0, " TOC \\o \"1-3\" \\h \\z \\u ", true);

        SdtBlock sdt = getTocSDT(wordMLPackage);

        System.out.println(sdt.getSdtContent().getContent().size());        
        System.out.println(wordMLPackage.getMainDocumentPart().getXML());

        Docx4J.save(wordMLPackage, new File("testToCHeadingEmpty.docx"));
    }

//    @Test
    public void testToCHeadingSet() throws TocException, Exception {

        Toc.setTocHeadingText("Alpha");

        WordprocessingMLPackage wordMLPackage = createPkg();

        TocGenerator tocGenerator = new TocGenerator(wordMLPackage);
        tocGenerator.generateToc( 0, " TOC \\o \"1-3\" \\h \\z \\u ", true);

        SdtBlock sdt = getTocSDT(wordMLPackage);

        System.out.println(sdt.getSdtContent().getContent().size());        
        System.out.println(wordMLPackage.getMainDocumentPart().getXML());

        Docx4J.save(wordMLPackage, new File("testToCHeadingSet.docx"));
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
