package vn.edu.likelion.entity;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import vn.edu.likelion.utility.Constant;

import java.io.*;
import java.util.Base64;

public class DocumentApp {
    private XWPFDocument document;

    public DocumentApp() {
        this.document = new XWPFDocument();
    }

    public void writeFileDocument(){
        String line;

        try {
            BufferedReader bufferedReader = new BufferedReader(new FileReader(Constant.SOURCE_FILE));
            while ((line = bufferedReader.readLine()) != null){
                String[] s = line.split("\t");
                if (s[2].equals("0")){
                    XWPFParagraph paragraph = document.createParagraph();
                    XWPFRun run = paragraph.createRun();

                    String encode = Base64.getEncoder().encodeToString(s[1].getBytes());

                    run.setText(s[0]);
                    run.addTab();
                    run.setText(encode);
                }
            }
        }catch (IOException e){
            System.out.println("Taking database fail!");
        }

        try {
            FileOutputStream stream = new FileOutputStream(Constant.FILE_ABSENT);
            document.write(stream);
            document.close();
            stream.close();
        } catch (IOException e) {
            System.out.println("Writing file word fail!");
        }
    }

    public void print(){
        File file = new File(Constant.FILE_ABSENT);

        try {
            InputStream fis = new FileInputStream(file);
            document = new XWPFDocument(OPCPackage.open(fis));
            fis.close();
        } catch (InvalidFormatException | IOException e) {
            e.printStackTrace();
        }

        System.out.println("List of student is absent");

        for (XWPFParagraph paragraph1 : document.getParagraphs()){
            String[] s = paragraph1.getText().split("\t");
            String id = s[0].trim();
            String encodeName = s[1].trim();
            String fullName = new String(Base64.getDecoder().decode(encodeName));
            System.out.println(id + "\t" + fullName);
        }

        try {
            document.close();
        } catch (IOException e) {
            e.printStackTrace();
        }

    }
}
