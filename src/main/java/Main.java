/**
 * Created by Moncef on 14/03/2016.
 */

import com.gargoylesoftware.htmlunit.*;
import com.gargoylesoftware.htmlunit.html.*;
import com.mongodb.Block;
import com.mongodb.MongoClient;
import com.mongodb.client.FindIterable;
import com.mongodb.client.MongoCollection;
import com.mongodb.client.MongoDatabase;
import org.apache.commons.logging.LogFactory;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.bson.Document;

import java.io.FileOutputStream;
import java.io.IOException;
import java.net.MalformedURLException;
import java.net.URL;
import java.util.ArrayList;
import java.util.List;
import java.util.concurrent.atomic.AtomicReference;
import java.util.logging.Level;
import java.util.logging.Logger;

class Main {
    /**
     * @param args
     */
    public static void main(String[] args) throws IOException {
        //extractIds();
        //extractStations();
        /**
         * Connect to Mongo DB server, get the database and the collection
         */
        MongoClient mongoClient = new MongoClient("localhost", 27017);
        MongoDatabase database = mongoClient.getDatabase("assainissement");
        MongoCollection<Document> stations = database.getCollection("stations");
        FindIterable<Document> all = stations.find();
        /**
         * initialize the docs list and get all docs
         */
        final List<Document> docs = new ArrayList<Document>();
        all.forEach(new Block<Document>() {
            @Override
            public void apply(final Document document) {
                docs.add(document);
            }
        });
        try {
            /**
             * execute writeXLSXFile with the list documents
             */
            long startTime = System.currentTimeMillis();
            writeXLSXFile(docs);
            System.out.println(System.currentTimeMillis()-startTime);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    /**
     * extract all stations
     */
    private static void extractStations() {
        MongoClient mongoClient = new MongoClient("localhost", 27017);
        MongoDatabase database = mongoClient.getDatabase("assainissement");
        MongoCollection<Document> stations = database.getCollection("stations");
        MongoCollection<Document> listStation = database.getCollection("listStations");
        final List<String> codes = new ArrayList<String>();
        FindIterable<Document> all = listStation.find().skip((int) stations.count() + 15);
        all.forEach(new Block<Document>() {
            @Override
            public void apply(final Document document) {
                codes.add((String) document.get("id"));
            }
        });
        for (String code : codes) {
            Document dc = new Document();
            String url = "http://assainissement.developpement-durable.gouv.fr/fiche.php?code=" + code;
            try {
                WebClient webClient = new WebClient(BrowserVersion.FIREFOX_24);
                webClient.getOptions().setJavaScriptEnabled(true);
                webClient.getOptions().setThrowExceptionOnScriptError(false);
                webClient
                        .setAjaxController(new NicelyResynchronizingAjaxController());
                WebRequest request = new WebRequest(
                        new URL(url));
                LogFactory.getFactory().setAttribute("org.apache.commons.logging.Log", "org.apache.commons.logging.impl.NoOpLog");
                java.util.logging.Logger.getLogger("com.gargoylesoftware.htmlunit").setLevel(Level.OFF);
                java.util.logging.Logger.getLogger("org.apache.commons.httpclient").setLevel(Level.OFF);
                HtmlPage page = webClient.getPage(request);
                int i = webClient.waitForBackgroundJavaScript(1000);
                while (i > 0) {
                    i = webClient.waitForBackgroundJavaScript(1000);
                    if (i <= 2) {
                        break;
                    }
                    synchronized (page) {
                        try {
                            page.wait(100);
                        } catch (InterruptedException e) {
                            e.printStackTrace();
                        }
                    }
                }
                webClient.getAjaxController().processSynchron(page, request, true);
                List<HtmlDivision> divs = (List<HtmlDivision>) page.getByXPath("//div[@id='contenu']");
                List<HtmlAnchor> closed = (List<HtmlAnchor>) page.getByXPath("//a[@class='closed']");
                int a = nbdiv1(closed);
                int b = nbdiv2(closed, a);
                DomNodeList<HtmlElement> one = divs.get(0).getElementsByTagName("p");
                if (one.get(0).asText().substring(one.get(0).asText().indexOf(":") + 1).trim().length() > 0) {
                    dc.append(one.get(0).asText().substring(0, one.get(0).asText().indexOf(":")).trim(), one.get(0).asText().substring(one.get(0).asText().indexOf(":") + 1, one.get(0).asText().indexOf("(")).trim());
                }
                for (int j = 1; j < 14; j++) {
                    if (one.get(j).asText().substring(one.get(j).asText().indexOf(":") + 1).trim().length() > 0) {
                        dc.append(one.get(j).asText().substring(0, one.get(j).asText().indexOf(":")).trim(), one.get(j).asText().substring(one.get(j).asText().indexOf(":") + 1).trim());
                    }
                }
                DomNodeList<HtmlElement> two = divs.get(1).getElementsByTagName("p");
                for (int j = 0; j < two.size() - 1; j++) {
                    if (two.get(j).asText().substring(two.get(j).asText().indexOf(":") + 1).trim().length() > 0) {
                        dc.append(two.get(j).asText().substring(0, two.get(j).asText().indexOf(":")).trim(), two.get(j).asText().substring(two.get(j).asText().indexOf(":") + 1).trim());
                    }
                }
                DomNodeList<HtmlElement> three = divs.get(2).getElementsByTagName("p");
                Document obj = new Document();
                Document ob = new Document();
                Document mon = new Document();
                for (int j = 0; j < 3; j++) {
                    if (three.get(j).asText().substring(three.get(j).asText().indexOf(":") + 1).trim().length() > 0) {
                        obj.append(three.get(j).asText().substring(0, three.get(j).asText().indexOf(":")).trim(), three.get(j).asText().substring(three.get(j).asText().indexOf(":") + 1).trim());
                    }
                }
                if (three.get(0).asText().substring(three.get(0).asText().indexOf(":") + 1).trim().length() > 0)
                    dc.append("2014", obj);
                for (int j = 0; j < a; j++) {
                    if (j < closed.size()) closed.get(j).click();
                    try {
                        for (int l = 4 * (j + 1); l < 4 * (j + 1) + 3; l++) {
                            if (three.get(l).asText().substring(three.get(l).asText().indexOf(":") + 1).trim().length() > 0) {
                                ob.append(three.get(l).asText().substring(0, three.get(l).asText().indexOf(":")).trim(), three.get(l).asText().substring(three.get(l).asText().indexOf(":") + 1).trim());
                            }
                        }
                        dc.append(String.valueOf(2014 - j - 1), ob);
                    } catch (IndexOutOfBoundsException e) {
                        dc.append(String.valueOf(2014 - j - 1), ob);
                    } catch (NullPointerException e) {
                        dc.append(String.valueOf(2014 - j - 1), ob);
                    }
                }
                DomNodeList<HtmlElement> four = divs.get(3).getElementsByTagName("p");
                for (int j = 0; j < 7; j++) {
                    if (four.get(j).asText().substring(four.get(j).asText().indexOf(":") + 1).trim().length() > 0) {
                        dc.append(four.get(j).asText().substring(0, four.get(j).asText().indexOf(":")).trim(), four.get(j).asText().substring(four.get(j).asText().indexOf(":") + 1).trim());
                    }
                }
                DomNodeList<HtmlElement> five = divs.get(4).getElementsByTagName("p");
                for (int j = 0; j < 3; j++) {
                    if (five.get(j).asText().substring(five.get(j).asText().indexOf(":") + 1).trim().length() > 0) {
                        mon.append(five.get(j).asText().substring(0, five.get(j).asText().indexOf(":")).trim(), five.get(j).asText().substring(five.get(j).asText().indexOf(":") + 1).trim());
                    }
                }
                if (five.get(0).asText().substring(five.get(0).asText().indexOf(":") + 1).trim().length() > 0)
                    dc.append("2014", obj);
                for (int j = a; j < a + b; j++) {
                    if (j < closed.size()) closed.get(j).click();
                    try {
                        for (int l = 3 + (2 * (j - a)); l < 5 + (2 * (j - a)); l++) {
                            if (five.get(l).asText().substring(five.get(l).asText().indexOf(":") + 1).trim().length() > 0) {
                                mon.append(five.get(l).asText().substring(0, five.get(l).asText().indexOf(":")).trim(), five.get(l).asText().substring(five.get(l).asText().indexOf(":") + 1).trim());
                            }
                        }
                        dc.append("Respect de la réglementation", mon);
                    } catch (IndexOutOfBoundsException e) {
                        dc.append(String.valueOf(2014 - j - 1), ob);
                    } catch (NullPointerException e) {
                        dc.append(String.valueOf(2014 - j - 1), ob);
                    }
                }
                stations.insertOne(dc);
            } catch (FailingHttpStatusCodeException e) {
                e.printStackTrace();
            } catch (MalformedURLException e) {
                e.printStackTrace();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
    }

    /**
     * exract the ods of all stations
     */
    private static void extractIds() {
        String url;
        url = "http://assainissement.developpement-durable.gouv.fr/liste.php";
        Logger logger = Logger.getLogger("com.steps");
        logger.setLevel(Level.INFO);
        MongoClient mongoClient = new MongoClient("localhost", 27017);
        MongoDatabase database = mongoClient.getDatabase("assainissement");
        MongoCollection<Document> collection = database.getCollection("listStations");
        try {
            WebClient webClient = new WebClient(BrowserVersion.FIREFOX_24);
            webClient.getOptions().setJavaScriptEnabled(true);
            webClient.getOptions().setThrowExceptionOnScriptError(false);
            webClient
                    .setAjaxController(new NicelyResynchronizingAjaxController());
            WebRequest request = new WebRequest(
                    new URL(url));
            LogFactory.getFactory().setAttribute("org.apache.commons.logging.Log", "org.apache.commons.logging.impl.NoOpLog");
            java.util.logging.Logger.getLogger("com.gargoylesoftware.htmlunit").setLevel(Level.OFF);
            java.util.logging.Logger.getLogger("org.apache.commons.httpclient").setLevel(Level.OFF);
            HtmlPage page = webClient.getPage(request);
            int i = webClient.waitForBackgroundJavaScript(4000);
            while (i > 0) {
                i = webClient.waitForBackgroundJavaScript(4000);
                if (i <= 2) {
                    break;
                }
                synchronized (page) {
                    try {
                        page.wait(500);
                    } catch (InterruptedException e) {
                        e.printStackTrace();
                    }
                }
            }
            webClient.getAjaxController().processSynchron(page, request, true);
            HtmlTable table = page.getHtmlElementById("example");
            int l = 0;
            for (final HtmlTableRow row : table.getRows()) {
                l++;
                Document dc = new Document();
                if (l > 2) {
                    for (int k = 2; k < row.getCells().size(); k = k + 8) {
                        if (k == 2) {
                            DomElement e = row.getCells().get(k);
                            dc.append("name", e.asText());
                        }
                    }
                    for (int k = 0; k < row.getCells().size(); k = k + 8) {
                        if (k == 8) {
                            DomElement e = row.getCells().get(k).getFirstElementChild();
                            dc.append("code", e.getAttribute("href").substring(17, 29));
                        }
                    }
                    collection.insertOne(dc);
                }
            }
            HtmlAnchor pagination = (HtmlAnchor) page.getElementById("example_next");
            pagination.click();
            for (int j = 2; j <= 1313; j++) {
                synchronized (page) {
                    try {
                        page.wait(500);
                    } catch (InterruptedException e) {
                        e.printStackTrace();
                    }
                }
                webClient.getAjaxController().processSynchron(page, request, true);
                table = page.getHtmlElementById("example");
                l = 0;
                for (final HtmlTableRow row : table.getRows()) {
                    l++;
                    Document dcc = new Document();
                    if (l > 2) {
                        for (int k = 2; k < row.getCells().size(); k = k + 8) {
                            if (k == 2) {
                                DomElement e = row.getCells().get(k);
                                dcc.append("name", e.asText());
                            }
                        }
                        for (int k = 0; k < row.getCells().size(); k = k + 8) {
                            if (k == 8) {
                                DomElement e = row.getCells().get(k).getFirstElementChild();
                                String s = e.getAttribute("href");
                                dcc.append("code", s.substring(s.indexOf("=") + 1, s.indexOf("#")));
                            }
                        }
                        collection.insertOne(dcc);
                    }
                }
                pagination.click();
            }
        } catch (FailingHttpStatusCodeException e) {
            e.printStackTrace();
        } catch (MalformedURLException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }

    }

    /**
     * calculate the number of <a> with class "closed" in the third div
     *
     * @param closed list of all html anchor with class "closed"
     * @return number of <a> with class "closed" in the third div
     */
    private static int nbdiv1(List<HtmlAnchor> closed) {
        int a = 0;
        for (int i = closed.size() - 1; i >= 0; i--) {
            if (closed.get(i).asText().contains("2007") && closed.get(i).asText().contains("clefs")) {
                a = 7;
                break;
            }
            if (closed.get(i).asText().contains("2008") && closed.get(i).asText().contains("clefs")) {
                a = 6;
                break;
            }
            if (closed.get(i).asText().contains("2009") && closed.get(i).asText().contains("clefs")) {
                a = 5;
                break;
            }
            if (closed.get(i).asText().contains("2010") && closed.get(i).asText().contains("clefs")) {
                a = 4;
                break;
            }
            if (closed.get(i).asText().contains("2011") && closed.get(i).asText().contains("clefs")) {
                a = 3;
                break;
            }
            if (closed.get(i).asText().contains("2012") && closed.get(i).asText().contains("clefs")) {
                a = 2;
                break;
            }
            if (closed.get(i).asText().contains("2013") && closed.get(i).asText().contains("clefs")) {
                a = 1;
                break;
            }
        }
        return a;
    }

    /**
     * calculate the number of <a> with class "closed" in the fifth div
     *
     * @param closed list of all html anchor with class "closed"
     * @param a      number of <a> with class "closed" in the third div
     * @return number of <a> with class "closed" in the fifth div
     */
    private static int nbdiv2(List<HtmlAnchor> closed, int a) {
        int b = 0;
        for (int i = closed.size() - 1; i >= a; i--) {
            if (closed.get(i).asText().contains("2007") && closed.get(i).asText().contains("Respect")) {
                b = 7;
                break;
            }
            if (closed.get(i).asText().contains("2008") && closed.get(i).asText().contains("Respect")) {
                b = 6;
                break;
            }
            if (closed.get(i).asText().contains("2009") && closed.get(i).asText().contains("Respect")) {
                b = 5;
                break;
            }
            if (closed.get(i).asText().contains("2010") && closed.get(i).asText().contains("Respect")) {
                b = 4;
                break;
            }
            if (closed.get(i).asText().contains("2011") && closed.get(i).asText().contains("Respect")) {
                b = 3;
                break;
            }
            if (closed.get(i).asText().contains("2012") && closed.get(i).asText().contains("Respect")) {
                b = 2;
                break;
            }
            if (closed.get(i).asText().contains("2013") && closed.get(i).asText().contains("Respect")) {
                b = 1;
                break;
            }
        }
        return b;
    }

    private static void writeXLSXFile(List<Document> docs) throws IOException {
        String excelFileName = "C:/Users/Moncef/Desktop/all.xlsx";
        String sheetName = "Sheet1";
        XSSFWorkbook wb = new XSSFWorkbook();
        XSSFSheet sheet = wb.createSheet(sheetName);
        /**
         * set the tab hedear
         */
        XSSFRow row = sheet.createRow(0);
        XSSFCell nomst = row.createCell(0);
        nomst.setCellValue("Nom de la station");
        XSSFCell codest = row.createCell(1);
        codest.setCellValue("Code de la station");
        XSSFCell naturest = row.createCell(2);
        naturest.setCellValue("Nature de la station");
        XSSFCell region = row.createCell(3);
        region.setCellValue("Région");
        XSSFCell departement = row.createCell(4);
        departement.setCellValue("Département");
        XSSFCell dateMiseEnService = row.createCell(5);
        dateMiseEnService.setCellValue("Date de mise en service");
        XSSFCell capNominal = row.createCell(6);
        capNominal.setCellValue("Capacité nominale");
        XSSFCell debitDeref = row.createCell(7);
        debitDeref.setCellValue("Débit de référence");
        XSSFCell codeAgg = row.createCell(8);
        codeAgg.setCellValue("Code de l'agglomération");
        XSSFCell nomAgg = row.createCell(9);
        nomAgg.setCellValue("Nom de l'agglomération");
        XSSFCell tailleAgg = row.createCell(10);
        tailleAgg.setCellValue("Taille de l'agglomération en 2014");
        XSSFCell bassinHydro = row.createCell(11);
        bassinHydro.setCellValue("Bassin hydrographique");
        XSSFCell type = row.createCell(12);
        type.setCellValue("Type");
        XSSFCell nom = row.createCell(13);
        nom.setCellValue("Nom");
        XSSFCell nomBV = row.createCell(14);
        nomBV.setCellValue("Nom du bassin versant");
        XSSFCell chMax14 = row.createCell(15);
        chMax14.setCellValue("2014 : Charge maximale en entrée");
        XSSFCell chDEM14 = row.createCell(16);
        chDEM14.setCellValue("2014 : Charge Débit entrant moyen");
        XSSFCell pB14 = row.createCell(17);
        pB14.setCellValue("2014 : Production de boues");
        XSSFCell cEnEq14 = row.createCell(18);
        cEnEq14.setCellValue("2014 : Conforme en équipement");
        XSSFCell cEnEPerf14 = row.createCell(19);
        cEnEPerf14.setCellValue("2014 : Conforme en performance");
        XSSFCell chMax13 = row.createCell(20);
        chMax13.setCellValue("2013 : Charge maximale en entrée");
        XSSFCell chDEM13 = row.createCell(21);
        chDEM13.setCellValue("2013 : Charge Débit entrant moyen");
        XSSFCell pB13 = row.createCell(22);
        pB13.setCellValue("2013 : Production de boues");
        XSSFCell cEnEq13 = row.createCell(23);
        cEnEq13.setCellValue("2013 : Conforme en équipement");
        XSSFCell cEnEPerf13 = row.createCell(24);
        cEnEPerf13.setCellValue("2013 : Conforme en performance");
        XSSFCell chMax12 = row.createCell(25);
        chMax12.setCellValue("2012 : Charge maximale en entrée");
        XSSFCell chDEM12 = row.createCell(26);
        chDEM12.setCellValue("2012 : Charge Débit entrant moyen");
        XSSFCell pB12 = row.createCell(27);
        pB12.setCellValue("2012 : Production de boues");
        XSSFCell cEnEq12 = row.createCell(28);
        cEnEq12.setCellValue("2012 : Conforme en équipement");
        XSSFCell cEnEPerf12 = row.createCell(29);
        cEnEPerf12.setCellValue("2012 : Conforme en performance");
        XSSFCell chMax11 = row.createCell(30);
        chMax11.setCellValue("2011 : Charge maximale en entrée");
        XSSFCell chDEM11 = row.createCell(31);
        chDEM11.setCellValue("2011 : Charge Débit entrant moyen");
        XSSFCell pB11 = row.createCell(32);
        pB11.setCellValue("2011 : Production de boues");
        XSSFCell cEnEq11 = row.createCell(33);
        cEnEq11.setCellValue("2011 : Conforme en équipement");
        XSSFCell cEnEPerf11 = row.createCell(34);
        cEnEPerf11.setCellValue("2011 : Conforme en performance");
        XSSFCell chMax10 = row.createCell(35);
        chMax10.setCellValue("2010 : Charge maximale en entrée");
        XSSFCell chDEM10 = row.createCell(36);
        chDEM10.setCellValue("2010 : Charge Débit entrant moyen");
        XSSFCell pB10 = row.createCell(37);
        pB10.setCellValue("2010 : Production de boues");
        XSSFCell cEnEq10 = row.createCell(38);
        cEnEq10.setCellValue("2010 : Conforme en équipement");
        XSSFCell cEnEPerf10 = row.createCell(39);
        cEnEPerf10.setCellValue("2010 : Conforme en performance");
        XSSFCell chMax09 = row.createCell(40);
        chMax09.setCellValue("2009 : Charge maximale en entrée");
        XSSFCell chDEM09 = row.createCell(27);
        chDEM09.setCellValue("2009 : Charge Débit entrant moyen");
        XSSFCell pB09 = row.createCell(41);
        pB09.setCellValue("2009 : Production de boues");
        XSSFCell cEnEq09 = row.createCell(42);
        cEnEq09.setCellValue("2009 : Conforme en équipement");
        XSSFCell cEnEPerf09 = row.createCell(43);
        cEnEPerf09.setCellValue("2009 : Conforme en performance");
        XSSFCell chMax08 = row.createCell(44);
        chMax08.setCellValue("2008 : Charge maximale en entrée");
        XSSFCell chDEM08 = row.createCell(45);
        chDEM08.setCellValue("2008 : Charge Débit entrant moyen");
        XSSFCell pB08 = row.createCell(46);
        pB08.setCellValue("2008 : Production de boues");
        XSSFCell cEnEq08 = row.createCell(47);
        cEnEq08.setCellValue("2008 : Conforme en équipement");
        XSSFCell cEnEPerf08 = row.createCell(48);
        cEnEPerf08.setCellValue("2008 : Conforme en performance");
        for (int r = 0; r < docs.size(); r++) {
            row = sheet.createRow(r + 1);
            for (int c = 0; c < 48; c++) {
                nomst = row.createCell(0);
                nomst.setCellValue((String) docs.get(r).get("Nom de la station"));
                codest = row.createCell(1);
                codest.setCellValue((String) docs.get(r).get("Code de la station"));
                naturest = row.createCell(2);
                naturest.setCellValue((String) docs.get(r).get("Nature de la station"));
                region = row.createCell(3);
                region.setCellValue((String) docs.get(r).get("Région"));
                departement = row.createCell(4);
                departement.setCellValue((String) docs.get(r).get("Département"));
                dateMiseEnService = row.createCell(5);
                dateMiseEnService.setCellValue((String) docs.get(r).get("Date de mise en service"));
                capNominal = row.createCell(6);
                capNominal.setCellValue(Integer.parseInt(docs.get(r).get("Capacité nominale").toString().replace("EH", "").trim()));
                debitDeref = row.createCell(7);
                debitDeref.setCellValue(Integer.parseInt(docs.get(r).get("Débit de référence").toString().replace("m3/j", "").trim()));
                codeAgg = row.createCell(8);
                codeAgg.setCellValue((String) docs.get(r).get("Code de l'agglomération"));
                nomAgg = row.createCell(9);
                nomAgg.setCellValue((String) docs.get(r).get("Nom de l'agglomération"));
                tailleAgg = row.createCell(10);
                tailleAgg.setCellValue(Integer.parseInt(docs.get(r).get("Taille de l'agglomération en 2014").toString().replace("EH", "").trim()));
                bassinHydro = row.createCell(11);
                bassinHydro.setCellValue((String) docs.get(r).get("Bassin hydrographique"));
                type = row.createCell(12);
                type.setCellValue((String) docs.get(r).get("Type"));
                nom = row.createCell(13);
                nom.setCellValue((String) docs.get(r).get("Nom"));
                nomBV = row.createCell(14);
                nomBV.setCellValue((String) docs.get(r).get("Nom du bassin versant"));
                AtomicReference<Document> res = new AtomicReference<Document>((Document) docs.get(r).get("Respect de la réglementation"));
                AtomicReference<Document> doc14 = new AtomicReference<Document>((Document) docs.get(r).get("2014"));
                //TODO see the NA problem
                //TODO see the prblem of values in others years
                //TODO remove the units and convert to Integer
                try {
                    chMax14 = row.createCell(15);
                    //if (null != doc14.get().get("Charge maximale en entrée"))
                    chMax14.setCellValue(Integer.parseInt(doc14.get().get("Charge maximale en entrée").toString().replace("EH", "").trim()));
                    //else chMax14.setCellValue("NA");
                    chDEM14 = row.createCell(16);

                    //if (null != doc14.get().get("Débit entrant moyen"))
                    chDEM14.setCellValue(Integer.parseInt(doc14.get().get("Débit entrant moyen").toString().replace("m3/j", "").trim()));
                    // else chDEM14.setCellValue("NA");
                    pB14 = row.createCell(17);
                    //if (null != doc14.get().get("Production de boues"))
                    pB14.setCellValue(Integer.parseInt(doc14.get().get("Production de boues").toString().replace("tMS/an", "").trim()));
                    // else pB14.setCellValue("NA");
                    cEnEq14 = row.createCell(18);
                    cEnEq14.setCellValue((String) res.get().get("Conforme en équipement au 31/12/2014"));
                    cEnEPerf14 = row.createCell(19);
                    cEnEPerf14.setCellValue((String) res.get().get("Conforme en performance en 2014"));
                } catch (NullPointerException e) {
                    cEnEq14 = row.createCell(18);
                    cEnEq14.setCellValue("NA");
                    cEnEPerf14 = row.createCell(19);
                    cEnEPerf14.setCellValue("NA");
                }
                Document doc13;
                doc13 = (Document) docs.get(r).get("2013");
                try {
                    chMax13 = row.createCell(20);
                    //if (null != doc13.get("Charge maximale en entrée"))
                    chMax13.setCellValue(Integer.parseInt(doc13.get("Charge maximale en entrée").toString().replace("EH", "").trim()));
                    // else chMax13.setCellValue("NA");
                    chDEM13 = row.createCell(21);
                    //if (null != doc13.get("Débit entrant moyen"))
                    chDEM13.setCellValue(Integer.parseInt(doc13.get("Débit entrant moyen").toString().replace("m3/j", "").trim()));
                    //else chDEM13.setCellValue("NA");
                    pB13 = row.createCell(22);
                    // if (null != doc13.get("Production de boues"))
                    pB13.setCellValue(Integer.parseInt(doc13.get("Production de boues").toString().replace("tMS/an", "").trim()));
                    //else pB13.setCellValue("NA");
                    cEnEq13 = row.createCell(23);
                    cEnEq13.setCellValue((String) res.get().get("Conforme en équipement au 31/12/2013"));
                    cEnEPerf13 = row.createCell(24);
                    cEnEPerf13.setCellValue((String) res.get().get("Conforme en performance en 2013"));
                } catch (NullPointerException e) {
                    cEnEq13 = row.createCell(23);
                    cEnEq13.setCellValue("NA");
                    cEnEPerf13 = row.createCell(24);
                    cEnEPerf13.setCellValue("NA");
                }
                chMax12 = row.createCell(25);
                Document doc12 = (Document) docs.get(r).get("2012");
                try {
                    chMax12.setCellValue(Integer.parseInt(doc12.get("Charge maximale en entrée").toString().replace("EH", "").trim()));
                    chDEM12 = row.createCell(26);
                    chDEM12.setCellValue(Integer.parseInt(doc12.get("Débit entrant moyen").toString().replace("m3/j", "").trim()));
                    pB12 = row.createCell(27);
                    pB12.setCellValue(Integer.parseInt(doc12.get("Production de boues").toString().replace("tMS/an", "").trim()));
                    cEnEq12 = row.createCell(28);
                    cEnEq12.setCellValue((String) res.get().get("Conforme en équipement au 31/12/2012"));
                    cEnEPerf12 = row.createCell(29);
                    cEnEPerf12.setCellValue((String) res.get().get("Conforme en performance en 2012"));
                } catch (NullPointerException e) {
                    cEnEq12 = row.createCell(28);
                    cEnEq12.setCellValue("NA");
                    cEnEPerf12 = row.createCell(29);
                    cEnEPerf12.setCellValue("NA");
                }
                chMax11 = row.createCell(30);
                Document doc11 = (Document) docs.get(r).get("2011");
                try {
                    chMax11.setCellValue(Integer.parseInt(doc11.get("Charge maximale en entrée").toString().replace("EH", "").trim()));
                    chDEM11 = row.createCell(31);
                    chDEM11.setCellValue(Integer.parseInt(doc11.get("Débit entrant moyen").toString().replace("m3/j", "").trim()));
                    pB11 = row.createCell(32);
                    pB11.setCellValue(Integer.parseInt(doc11.get("Production de boues").toString().replace("tMS/an", "").trim()));
                    cEnEq11 = row.createCell(33);
                    cEnEq11.setCellValue((String) res.get().get("Conforme en équipement au 31/12/2011"));
                    cEnEPerf11 = row.createCell(34);
                    cEnEPerf11.setCellValue((String) res.get().get("Conforme en performance en 2011"));
                } catch (NullPointerException e) {
                    cEnEq11 = row.createCell(33);
                    cEnEq11.setCellValue("NA");
                    cEnEPerf11 = row.createCell(35);
                    cEnEPerf11.setCellValue("NA");
                }
                Document doc10 = (Document) docs.get(r).get("2010");
                try {
                    chMax10 = row.createCell(36);
                    chMax10.setCellValue(Integer.parseInt(doc10.get("Charge maximale en entrée").toString().replace("EH", "").trim()));
                    chDEM10 = row.createCell(37);
                    chDEM10.setCellValue(Integer.parseInt(doc10.get("Débit entrant moyen").toString().replace("m3/j", "").trim()));
                    pB10 = row.createCell(38);
                    pB10.setCellValue(Integer.parseInt(doc10.get("Production de boues").toString().replace("tMS/an", "").trim()));
                    cEnEq10 = row.createCell(39);
                    cEnEq10.setCellValue((String) res.get().get("Conforme en équipement au 31/12/2010"));
                    cEnEPerf10 = row.createCell(40);
                    cEnEPerf10.setCellValue((String) res.get().get("Conforme en performance en 2010"));
                } catch (NullPointerException e) {
                    cEnEq10 = row.createCell(39);
                    cEnEq10.setCellValue("NA");
                    cEnEPerf10 = row.createCell(40);
                    cEnEPerf10.setCellValue("NA");
                }
                Document doc09 = (Document) docs.get(r).get("2009");
                try {
                    chMax09 = row.createCell(41);
                    chMax09.setCellValue(Integer.parseInt(doc09.get("Charge maximale en entrée").toString().replace("EH", "").trim()));
                    chDEM09 = row.createCell(42);
                    chDEM09.setCellValue(Integer.parseInt(doc09.get("Débit entrant moyen").toString().replace("m3/j", "").trim()));
                    pB09 = row.createCell(43);
                    pB09.setCellValue(Integer.parseInt(doc09.get("Production de boues").toString().replace("tMS/an", "").trim()));
                    cEnEq09 = row.createCell(44);
                    cEnEq09.setCellValue((String) res.get().get("Conforme en équipement au 31/12/2009"));
                    cEnEPerf09 = row.createCell(45);
                    cEnEPerf09.setCellValue((String) res.get().get("Conforme en performance en 2009"));
                } catch (NullPointerException e) {
                    cEnEq09 = row.createCell(44);
                    cEnEq09.setCellValue("NA");
                    cEnEPerf09 = row.createCell(45);
                    cEnEPerf09.setCellValue("NA");
                }
                Document doc08 = (Document) docs.get(r).get("2008");
                try {
                    chMax08 = row.createCell(46);
                    chMax08.setCellValue(Integer.parseInt(doc08.get("Charge maximale en entrée").toString().replace("EH", "").trim()));
                    chDEM08 = row.createCell(47);
                    chDEM08.setCellValue(Integer.parseInt(doc08.get("Débit entrant moyen").toString().replace("m3/j", "").trim()));
                    pB08 = row.createCell(48);
                    pB08.setCellValue(Integer.parseInt(doc08.get("Production de boues").toString().replace("tMS/an", "").trim()));
                    cEnEq08 = row.createCell(49);
                    cEnEq08.setCellValue((String) res.get().get("Conforme en équipement au 31/12/2008"));
                    cEnEPerf08 = row.createCell(50);
                    cEnEPerf08.setCellValue((String) res.get().get("Conforme en performance en 2008"));
                } catch (NullPointerException e) {
                    cEnEq08 = row.createCell(49);
                    cEnEq08.setCellValue("NA");
                    cEnEPerf08 = row.createCell(50);
                    cEnEPerf08.setCellValue("NA");
                }
            }
        }
        FileOutputStream fileOut = new FileOutputStream(excelFileName);
        wb.write(fileOut);
        fileOut.flush();
        fileOut.close();
    }
}