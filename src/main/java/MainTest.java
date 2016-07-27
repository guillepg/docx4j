import java.io.File;
import org.docx4j.TraversalUtil;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.docx4j.toc.TocFinder;
import org.docx4j.toc.TocGenerator;
import org.docx4j.wml.Body;
import org.docx4j.wml.Document;
import org.docx4j.wml.SdtBlock;

/**
 *
 * @author a940
 */
public class MainTest {
    
    final static String PATH = "C:/Users/a940/Desktop/CSCAE/";
    
    static String inputfilepath = PATH+"antes.docx";
    static String outputfilepath = PATH+"despues.docx";    
    static String inputLectura = PATH+"contabla.docx";	
    public static final String TOC_STYLE_MASK = "TOC%s";
    
    public static void main(String[] args){
        try{
            leerTituloDocumento(inputLectura);
            System.out.println("----------");
            leerDocumentoCompleto(inputLectura);
            System.out.println("----------");
            leerTablaContenidos(inputLectura);
        }catch(Exception ex){
            ex.printStackTrace();
        }
    }
    
    public static void leerTituloDocumento(String ruta) throws Docx4JException{
        WordprocessingMLPackage WMLP = WordprocessingMLPackage.load(new File(ruta));
        System.out.println("title: "+WMLP.getTitle());
    }
    
    public static void leerTablaContenidos(String ruta) throws Docx4JException{
        WordprocessingMLPackage WMLP = WordprocessingMLPackage.load(new File(ruta));
        SdtBlock tabla = (SdtBlock) WMLP.getMainDocumentPart().getContent().get(1);
        System.out.println(tabla.getSdtContent().getContent());
//        SdtBlock sdt = getTocSDT(WMLP);
//        TocFinder tocf = new TocFinder();
//        System.out.println("tocSDT: "+tocf.getTocSDT().toString());
    }
    
    public static void leerDocumentoCompleto(String ruta) throws Docx4JException{
        WordprocessingMLPackage WMLP = WordprocessingMLPackage.load(new File(ruta));
        System.out.println(WMLP.getMainDocumentPart().getXML());
    }
    
    public static void escribirToc1() throws Docx4JException{

        WordprocessingMLPackage WMLP = WordprocessingMLPackage.load(new File(inputfilepath));
        TocGenerator tocGenerator = new TocGenerator(WMLP);
        tocGenerator.generateToc(0,"TOC \\h \\z \\t \"comh1,1,comh2,2,comh3,3,comh4,4\" ", true);
//        tocGenerator.generateToc(0,"TOC \\o \"1-3\" \\h \\z \\u ", true);
        WMLP.save(new java.io.File(outputfilepath) );

    }
    
    public static void escribirToc2() throws Docx4JException, Exception{
        WordprocessingMLPackage WMLP = createPkg();
		
        TocGenerator tocGenerator = new TocGenerator(WMLP);		
        tocGenerator.generateToc( 0, " TOC \\o \"1-3\" \\h \\z \\u ", true);
//        SdtBlock sdt = getTocSDT(MLP);
//        System.out.println(sdt.getSdtContent().getContent().size());  
    }
    
    //<editor-fold defaultstate="collapsed" desc="METODOS AUXILIARES">    
    private static WordprocessingMLPackage createPkg() throws Exception{
    	
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
    
//    private static SdtBlock getTocSDT(WordprocessingMLPackage wordMLPackage) {
//
//        MainDocumentPart documentPart = wordMLPackage.getMainDocumentPart();
//        Document wmlDocumentEl = (Document)documentPart.getJaxbElement();
//        Body body = wmlDocumentEl.getBody();
//
//    	TocFinder finder = new TocFinder();
//        new TraversalUtil(body.getContent(), finder);
//
//        return finder.tocSDT;
//    }
    //</editor-fold>
}
