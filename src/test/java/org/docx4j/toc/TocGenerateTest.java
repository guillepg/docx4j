package org.docx4j.toc;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.PrintWriter;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.List;
import java.util.logging.Level;
import java.util.logging.Logger;
import javax.xml.bind.JAXBElement;
import junit.framework.Assert;
import org.junit.Test;

import org.docx4j.Docx4J;
import org.docx4j.TraversalUtil;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.docx4j.wml.Body;
import org.docx4j.wml.Document;
import org.docx4j.wml.SdtBlock;
import org.docx4j.wml.P;
import org.docx4j.wml.R;
import org.docx4j.openpackaging.parts.Part;
import org.docx4j.relationships.Relationship;
import org.docx4j.wml.Style;

public class TocGenerateTest {

    static final String TOC_STYLE_MASK = "TOC%s";
    final static String PATH = "C:/Users/a940/Desktop/";
    
    static String input = PATH + "fichero prueba parrafos.docx";
    static String inputcontabla = PATH + "CSCAE/tabla_con.docx";
    static String inputsintabla = PATH + "CSCAE/tabla_sin.docx";
    static String inputEstilos = PATH + "CSCAE/styles.DOCX";
    static String inputCSCAE = PATH + "Huesca.docx";
    
    static String outputXML1 = PATH + "estilos1.xml"; //CSCAE/txt/
    static String outputXML2 = PATH + "estilos2.xml";    
    static String outputDOCX = PATH + "output.docx";

    static final String FOOTER_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer";
    
    @Test // MODIFICABLE PARA LEER DISTINTAS PARTES
    public void testLecturaPartesXML(){
        try {
            // Lectura de MainDocument y escritura en XML
            WordprocessingMLPackage WMLP1 = WordprocessingMLPackage.load(new File(inputEstilos));
            System.out.println("Reading " + WMLP1.name());                        
            String stylesXML1 = WMLP1.getMainDocumentPart().getStyleDefinitionsPart(false).getXML();
            PrintWriter pw1 = new PrintWriter(outputXML1);
            pw1.write(stylesXML1);
            pw1.close();
            
            WordprocessingMLPackage WMLP2 = WordprocessingMLPackage.load(new File(inputcontabla));
            System.out.println("Reading " + WMLP2.name());            
            String stylesXML2 = WMLP2.getMainDocumentPart().getStyleDefinitionsPart(false).getXML();
            PrintWriter pw2 = new PrintWriter(outputXML2);
            pw2.write(stylesXML2);
            pw2.close();

            // Lectura de DocumentSettings
//            String settingsXML = WMLP.getMainDocumentPart().getDocumentSettingsPart().getXML();
//            if(settingsXML.contains("updateFields")){/*comprobar si es true*/}
//            else{/*añadir campo en linea ultima-1*/}   
            
            // Lectura de Table of Content
//            SdtBlock sdt = getTocSDT(WMLP);
//            leerBloqueSdt(sdt);
            
        } catch (Docx4JException docxex) {
            Logger.getLogger(TocGenerateTest.class.getName()).log(Level.SEVERE, "ERROR DE DOCX4J", docxex);
        } catch (FileNotFoundException fnfex) {
            Logger.getLogger(TocGenerateTest.class.getName()).log(Level.SEVERE, "ERROR DE LECTURA EN FICHERO", fnfex);
        } catch (Exception ex) {
            Logger.getLogger(TocGenerateTest.class.getName()).log(Level.SEVERE, "EXCEPCION", ex);
        }
    }
   
//    @Test // WORKING FINE
    public void testCreacionAuto(){
        try {
            WordprocessingMLPackage WMLP = WordprocessingMLPackage.load(new File(inputCSCAE));
            TocGenerator tocGen = new TocGenerator(WMLP);
            tocGen.generateToc(0, "TOC \\o \"1-4\" \\h \\z \\u ", false);
            WMLP.save(new File(outputDOCX));
            Logger.getLogger(TocGenerateTest.class.getName()).log(Level.INFO, "testCreacionAuto OK!", "");
        } catch (Exception ex) {
            Logger.getLogger(TocGenerateTest.class.getName()).log(Level.SEVERE, "ERROR EN CREACION", ex);
        }
    }
    
//    @Test // WORKING FINE
    public void testActualizacion(){
        try {
            WordprocessingMLPackage WMLP = WordprocessingMLPackage.load(new File(inputcontabla));
            TocGenerator tocGen = new TocGenerator(WMLP);
            tocGen.updateToc(false);
            WMLP.save(new File(outputDOCX));
            Logger.getLogger(TocGenerateTest.class.getName()).log(Level.INFO, "testPropioActualizacion OK!", "");

        } catch (Exception ex) {
            Logger.getLogger(TocGenerateTest.class.getName()).log(Level.SEVERE, "ERROR EN ACTUALIZACION", ex);
        }
    } 
     
//    @Test // NOT USEFUL
    public void testCreacionManual(){
        try {
            // 0 -> Titulo
            // 1 -> Tabulacion (default: 8494)
            // 2 -> TOC \\o "1-3" \\h \\z \\u
//            WordprocessingMLPackage WMLP = WordprocessingMLPackage.load(new File(inputsintabla));
//            String text = new String(Files.readAllBytes(Paths.get(plantilla)), StandardCharsets.UTF_8);
//            
//            text = text.replace("{0}","titulo");
//            text = text.replace("{1}","8494");
//            text = text.replace("{2}","TOC \\o \"1-3\" \\h \\z \\u");
//            System.out.println(text);
//            
//            Logger.getLogger(TocGenerateTest.class.getName()).log(Level.INFO, "Escritura OK", "");
        } catch (Exception ex) {
            Logger.getLogger(TocGenerateTest.class.getName()).log(Level.SEVERE, "ERROR EN ESCRITURA", ex);
        }
    }
    
    
    //<editor-fold defaultstate="collapsed" desc="codigo anterior">
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

    /* --------------------- MÉTODOS AUXILIARES --------------------- */
    private void leerBloqueSdt(SdtBlock sdt){
        int size = sdt.getSdtContent().getContent().size();
        System.out.println("leerBloqueSdt size: " + size);

        for(int i=0;i<size;i++){
            P pe = (P) sdt.getSdtContent().getContent().get(i);
            int subsize = pe.getContent().size();
            System.out.println("testPropioLectura elem "+i+": " + pe.getContent());
            System.out.println("testPropioLectura elem "+i+": size " + pe.getContent().size());

            for(int j=0;j<subsize;j++){
                if(pe.getContent().get(j).getClass().equals(R.class)){
                    R run = (R) pe.getContent().get(j);
                    System.out.println("subelem R"+i+"."+j+": "+run.getRsidR()+" - "+run.getRsidRPr());
                }
                else if(pe.getContent().get(j).getClass().equals(JAXBElement.class)){
                    JAXBElement jaxb = (JAXBElement) pe.getContent().get(j);
                    System.out.println("subelem JAXB "+i+"."+j+": " + jaxb.getValue());
                }
            }
        }
    }
            
//    for(int i=0;i<tabla.getSdtContent().getContent().size();i++){
//        try {
//            P pe = (P) tabla.getSdtContent().getContent().get(i);
//            JAXBElement jaxb = (JAXBElement) pe.getContent().get(0);
//            Hyperlink link = (Hyperlink) jaxb.getValue();
//            R erre = (R) link.getContent().get(0);
//            Text texto = (Text) erre.getContent().get(0);
//            texto.setValue(texto.getValue()+" NUEVO");
//            System.out.println("Texto: "+ texto.getValue());
//        } 
//        catch (Exception ex) {}                
//    }
    
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
    //</editor-fold>
}
