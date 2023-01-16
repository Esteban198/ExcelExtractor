import java.io.*;
import java.net.HttpURLConnection;
import java.net.URL;
import java.net.http.HttpClient;
import java.net.http.HttpRequest;

import org.apache.tika.exception.TikaException;
import org.apache.tika.metadata.Metadata;
import org.apache.tika.parser.ParseContext;
import org.apache.tika.parser.microsoft.ooxml.OOXMLParser;
import org.apache.tika.sax.BodyContentHandler;

import org.xml.sax.SAXException;


public class Main {

    public static void main(final String[] args) throws IOException, TikaException, SAXException {

        // detecting the file type
        BodyContentHandler handler = new BodyContentHandler();
        Metadata metadata = new Metadata();
        FileInputStream inputstream = new FileInputStream(new File("C:\\solr-8.11.2\\solr-8.11.2\\server\\solr\\test\\organizations-1000test.xlsx"));
        ParseContext pcontext = new ParseContext();

        // OOXml parser
        OOXMLParser msofficeparser = new OOXMLParser();
        msofficeparser.parse(inputstream, handler, metadata, pcontext);
        System.out.println(handler);
        System.out.println("Contenido del Excel: " + handler.toString().replaceAll("\t", ","));
        System.out.println("Metadata del excel: ");
        String[] metadataNames = metadata.names();

        for (String name : metadataNames) {
             System.out.println(name + ": " + metadata.get(name));
        }

        File csvFile = new File("archivo.csv");
        FileWriter fileWriter = new FileWriter(csvFile);
        String[] post = handler.toString().replaceAll(",",";" ).split("\t");
        post[0]="";
        String dataPost = "";
        for (String data : post) {
            dataPost += data +",";
        }
        System.out.println(dataPost);
        dataPost= dataPost.replaceAll(".$", "");
        fileWriter.write(dataPost);

        URL url = new URL("http://localhost:8983/solr/ejemplo/update?commit=true");
        HttpURLConnection connection = (HttpURLConnection) url.openConnection();
        connection.setDoOutput(true);
        connection.setInstanceFollowRedirects(false);
        connection.setRequestMethod("POST");
        connection.setRequestProperty("Content-Type", "text/csv");
        connection.connect();

        connection.getOutputStream().write(dataPost.getBytes());

        System.out.println(connection.getResponseMessage());

    }
}