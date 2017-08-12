package WordP;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.math.BigInteger;
import javax.swing.*;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.util.Units;
import org.apache.poi.xwpf.extractor.XWPFWordExtractor;
import org.apache.poi.xwpf.usermodel.*;
import org.apache.poi.xwpf.usermodel.XWPFTable.XWPFBorderType;
import org.apache.poi.xwpf.usermodel.XWPFTableCell.XWPFVertAlign;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTblBorders;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTblPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STBorder;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STMerge;

/**
 * Created by Overlord on 15.07.2017.
 */
public class wordp {

    private static void mergeCellsVertically(XWPFTable table, int col, int      fromRow, int toRow) {
        for (int rowIndex = fromRow; rowIndex <= toRow; rowIndex++) {
            XWPFTableCell cell = table.getRow(rowIndex).getCell(col);
            if ( rowIndex == fromRow ) {
                // The first merged cell is set with RESTART merge value
                cell.getCTTc().addNewTcPr().addNewVMerge().setVal(STMerge.RESTART);
            } else {
                // Cells which join (merge) the first one, are set with CONTINUE
                cell.getCTTc().addNewTcPr().addNewVMerge().setVal(STMerge.CONTINUE);
            }
        }
    }

    private static void mergeCellsHorizontally(XWPFTable table, int row, int fromCol, int toCol) {

        for (int colIndex = fromCol; colIndex <= toCol; colIndex++) {

            XWPFTableCell cell = table.getRow(row).getCell(colIndex);

            if ( colIndex == fromCol ) {
                // The first merged cell is set with RESTART merge value
                cell.getCTTc().addNewTcPr().addNewHMerge().setVal(STMerge.RESTART);
            } else {
                // Cells which join (merge) the first one, are set with CONTINUE
                cell.getCTTc().addNewTcPr().addNewHMerge().setVal(STMerge.CONTINUE);
            }
        }
    }

    public static void main(String[] args) throws FileNotFoundException, IOException, InvalidFormatException {
        // TODO code application logic here
        XWPFDocument document =new XWPFDocument();
        XWPFParagraph paragraph = document.createParagraph();
        XWPFRun run = paragraph.createRun();
        paragraph.setSpacingAfter(0);
        String imgFile = "logo.jpg";
        FileInputStream is = new FileInputStream(imgFile);
       // run.addBreak();
        run.addPicture(is, XWPFDocument.PICTURE_TYPE_JPEG, imgFile, Units.toEMU(48), Units.toEMU(48)); // 200x200 pixels
        //run.addBreak();
        is.close();

       // run.setText("xcgsdg");

        //run.setText("Title");
        //run.setFontSize(25);
        //run.setFontFamily("Arial Black");
        //run.setUnderline(UnderlinePatterns.DOUBLE);
        //run.setBold(true);
        //run.setItalic(true);
        //run.setStrike(true);
        //run.setColor("009E2F");
        paragraph.setAlignment(ParagraphAlignment.CENTER);
        //paragraph.setIndentationHanging(1000);//otstup?
        //paragraph.setBorderBottom(Borders.BASIC_THIN_LINES);//Linii
        //paragraph.setBorderTop(Borders.BASIC_THIN_LINES);
        //paragraph.setSpacingAfter(1000);
       // paragraph.setNumID(BigInteger.ONE);

     //   XWPFParagraph paragraph2=document.createParagraph();
        //paragraph2.setNumID(BigInteger.ONE);
        //paragraph2.setPageBreak(true);//razdelenie stranic
      //  XWPFRun run2=paragraph2.createRun();
      //  run2.setText("sdfsdfsd");

        //TABLICA
        XWPFTable tab = document.createTable();
        CTTblPr tblpro= tab.getCTTbl().getTblPr();
        CTTblBorders borders=tblpro.addNewTblBorders();
        borders.addNewBottom().setVal(STBorder.NONE);
        borders.addNewLeft().setVal(STBorder.NONE);
        borders.addNewRight().setVal(STBorder.NONE);
        borders.addNewTop().setVal(STBorder.NONE);
        tab.setInsideHBorder(XWPFBorderType.NONE,1,1,"000000");//, 3, 2, "0000BB"
        tab.setInsideVBorder(XWPFBorderType.NONE,1,1,"000000");
       // tab.setCellMargins(2,2,2,2);
        /*tab.setStyleID("1");
        String buf=tab.getStyleID();
        System.out.println(buf);*/
        //tab.setCellMargins(200,200,200,200);


        XWPFTableRow row0= tab.getRow(0);
        //row0.setHeight(500);

        XWPFTableCell cell0=row0.getCell(0);


        cell0.setVerticalAlignment(XWPFVertAlign.CENTER);
        //cell0.setColor("009E2F");

        XWPFParagraph par=cell0.addParagraph();
        par.setSpacingAfter(0);
        XWPFRun r=par.createRun();
        r.setText("Упраўленне сельскай гаспадаркі і харчавання\n" +
                "Рэчыцкага райвыканкама\n" +
               // "\n" +
                "КАМУНАЛЬНАЕ\n" +
                "СЕЛЬСКАГАСПАДАРЧАЕ УНІТАРНАЕ\n" +
                "ПРАДПРЫЕМСТВА\n" +
                "«АГРАКАМБІНАТ «ХОЛМЕЧ»\n" +
                "вул. Маладзежная, 14, 247505, в.Холмеч,\n" +
                "Рэчыцкі раён, Гомельская  вобласць\n" +
                "тэл.(02340) 33889, тэл./факс 33896\n" +
                "holmechagro@tut.by\n" +
                "р/р 3012200390014 у адз. ААО «Белаграпромбанк»\n" +
                "г. Рэчыца, код банка 940, УНП 400000006");
        par.setAlignment(ParagraphAlignment.CENTER);
        cell0.removeParagraph(0);
        cell0.setParagraph(par);
        // cell0.setParagraph(paragraph);

        XWPFTableCell cellBuf=row0.createCell();
        cellBuf.setText("  ");

        XWPFTableCell cell_1=row0.createCell();
        XWPFParagraph par1=cell_1.addParagraph();
        XWPFRun r1=par1.createRun();
        String buf="Управление сельского хозяйства и продовольствия\n" +
                "Речицкого райисполкома\n" +
                //   "\n" +
                //  "Речицкого райисполкома\n" +
                "КОММУНАЛЬНОЕ\n" +
                "СЕЛЬСКОХОЗЯЙСТВЕННОЕ УНИТАРНОЕ\n" +
                "ПРЕДПРИЯТИЕ\n" +
                "«АГРОКОМБИНАТ «ХОЛМЕЧ»\n" +
                "ул. Молодежная, 14, 247505, д.Холмеч,\n" +
                "Речицкий район, Гомельская обласць\n" +
                "тэл.(02340) 33889, тэл./факс 33896\n" +
                "holmechagro@tut.by\n" +
                "р/с3012200390014 в отд. ААО «Белаграпромбанк»\n" +
                "г. Речица  код банка 940, УНП 400000006";
        r1.setText(buf.toString());
        r1.addBreak();
        r1.setText("dfhgdf",1);
        par1.setAlignment(ParagraphAlignment.CENTER);
        par1.setSpacingAfter(0);
        cell_1.removeParagraph(0);
        cell_1.setParagraph(par1);
        cell_1.setVerticalAlignment(XWPFVertAlign.CENTER);

        XWPFTableRow rowDate=tab.createRow();
       /* row_1.removeCell(1);
        row_1.removeCell(0);*/
        XWPFTableCell cell1_0=rowDate.getCell(0);
        cell1_0.setVerticalAlignment(XWPFVertAlign.TOP);
        cell1_0.setText("Дата");
        //mergeCellsHorizontally(tab,1,0,1);


        XWPFTableRow rowAdr=tab.createRow();
        XWPFTableCell celladr=rowAdr.getCell(2);
        celladr.setVerticalAlignment(XWPFVertAlign.TOP);
        celladr.setText("Adresat");
        //mergeCellsHorizontally(tab,2,1,2);

        //ОБъединить ячейки


       // mergeCellsVertically(tab,0,0,1);











       /* XWPFTableCell cell1_1=row_1.getCell(1);
        cell1_1.setText("NewLine2");*/

        /*XWPFTableCell cell1_2=row_1.createCell();
        cell1_2.setText("NewLine3");*/

        //GeneratorTexta
       /* XWPFParagraph content=document.createParagraph();
        XWPFRun contentRun=content.createRun();

        for(int numberOfWords=0; numberOfWords<1000; numberOfWords++){
            StringBuilder sb = new StringBuilder();

            for(int numberOfChars=0; numberOfChars<5; numberOfChars++){
                char c= (char)(Math.random()*33+'а');
                sb.append(c);
            }
            sb.append(' ');
            contentRun.setText(sb.toString());
        }*/

     /* //Chtenie iz faila
        JFileChooser window=new JFileChooser();
        int returnValue = window.showOpenDialog(null);
        XWPFWordExtractor extract=null;
        if(returnValue==JFileChooser.APPROVE_OPTION){
            XWPFDocument doc= new XWPFDocument(new FileInputStream(window.getSelectedFile()));
             extract= new XWPFWordExtractor(doc);
            System.out.print(extract.getText());
        }*/





       /* XWPFParagraph paragraph3=document.createParagraph();
        XWPFRun r3=paragraph3.createRun();
        r3.setText("!!!!"+extract.getText());*/



        try{
            FileOutputStream output = new FileOutputStream("Proba.docx");
            document.write(output);
            output.close();
        }catch(Exception e){
            e.printStackTrace();
        }

    }

}
