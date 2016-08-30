package org.docx4j.toc;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.PrintWriter;
import java.math.BigInteger;
import java.util.List;
import java.util.logging.Level;
import java.util.logging.Logger;
import javax.xml.bind.JAXBElement;
import junit.framework.Assert;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xwpf.usermodel.XWPFAbstractNum;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.junit.Test;

import org.docx4j.Docx4J;
import org.docx4j.TraversalUtil;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.apache.poi.xwpf.usermodel.XWPFNumbering;
import org.docx4j.wml.Body;
import org.docx4j.wml.Document;
import org.docx4j.wml.SdtBlock;
import org.docx4j.wml.P;
import org.docx4j.wml.R;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTLvl;
//import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTMultiLevelType;

public class TocGenerateTest {

    static final String TOC_STYLE_MASK = "TOC%s";
    final static String PATH = "C:/Users/a940/Desktop/CSCAE/";
    
//    static String inputParrafos = PATH + "ejemplos/fichero prueba parrafos.docx";
//    static String inputConTabla = PATH + "ejemplos/tabla_con.docx";
//    static String inputSinTabla = PATH + "ejemplos/tabla_sin.docx";
    static String inputEstilos = PATH + "ejemplos/styles.DOCX";
    static String inputConLista = PATH + "ejemplos/listaejemplo.docx";
    static String inputCSCAE = PATH + "Rioja.docx";
    
    static String outputXML1 = PATH + "salida1.xml";
    static String outputXML2 = PATH + "salida2.xml";    
    static String outputDOCX = PATH + "output.docx";

    static final String FOOTER_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer";
    
    @Test // MODIFICABLE PARA LEER DISTINTAS PARTES
    public void testLecturaPartesXML(){
        try {
            XWPFDocument doc = new XWPFDocument(XWPFDocument.openPackage(inputConLista));
            XWPFNumbering numbering = doc.getNumbering();
            PrintWriter pw1 = new PrintWriter(outputXML1);
            if(numbering==null){
                System.out.println(" ---> No hay numeración <--- "); 
//                numbering = doc.createNumbering();
//                doc.write(new FileOutputStream(outputDOCX));
            }
            else{
                for(int i=0;i<10;i++){
                    XWPFAbstractNum xwpfabsnum = numbering.getAbstractNum(BigInteger.valueOf(i));
                    boolean copiar = false;
                    if(xwpfabsnum!=null){
                        pw1.write("Numbering "+i+": ");
                        CTLvl[] array = xwpfabsnum.getAbstractNum().getLvlArray();
                        for(int j=0;j<array.length;j++){
                            pw1.write("\n\t"+array[j].getNumFmt().getVal().toString());
                            pw1.write("\n\t"+array[j].getLvlText().getVal());
                            if(array[j].getLvlText().getVal().equals("%1.%2.")){
                                copiar = true;
                            }
                        }
                        numbering.addAbstractNum(xwpfabsnum);
                    }
                    pw1.write("\n");
                }
                // Lectura de MainDocument y escritura en XML
                WordprocessingMLPackage WMLP1 = WordprocessingMLPackage.load(new File(inputConLista));
                System.out.println("Reading " + WMLP1.name());
//                pw1.write("---------------------------ESTILOS-----------------------------\n");
//                pw1.write(WMLP1.getMainDocumentPart().getStyleDefinitionsPart(false).getXML());
                pw1.write("-------------------------NUMERACION----------------------------\n");
                if(WMLP1.getMainDocumentPart().getNumberingDefinitionsPart() != null) 
                    pw1.write(WMLP1.getMainDocumentPart().getNumberingDefinitionsPart().getXML());
                else 
                    pw1.write("No hay numeración");
//                pw1.write("-----------------------DocumentSettings------------------------\n");
//                pw1.write(WMLP1.getMainDocumentPart().getDocumentSettingsPart().getXML());
                pw1.write("\n------------------------DOCUMENTO----------------------------\n");
                pw1.write(WMLP1.getMainDocumentPart().getXML());
                pw1.close();
            }
               
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
//            WordprocessingMLPackage WMLP = WordprocessingMLPackage.load(new File(inputConTabla));
//            TocGenerator tocGen = new TocGenerator(WMLP);
//            tocGen.updateToc(false);
//            WMLP.save(new File(outputDOCX));
//            Logger.getLogger(TocGenerateTest.class.getName()).log(Level.INFO, "testPropioActualizacion OK!", "");

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
