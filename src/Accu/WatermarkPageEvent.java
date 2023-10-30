package Accu;

import com.itextpdf.text.Document;
import com.itextpdf.text.Element;
import com.itextpdf.text.Font;
import com.itextpdf.text.Phrase;
import com.itextpdf.text.pdf.ColumnText;
import com.itextpdf.text.pdf.GrayColor;
import com.itextpdf.text.pdf.PdfPageEventHelper;
import com.itextpdf.text.pdf.PdfWriter;

/**
 * Created by gill14 on 2/18/19.
 */
public class WatermarkPageEvent extends PdfPageEventHelper {
    Font FONT = new Font(Font.FontFamily.HELVETICA, 52, Font.BOLD, new GrayColor(0.20f));

    @Override
    public void onEndPage(PdfWriter writer, Document document) {
        if (writer.getPageNumber()==2) {
            ColumnText.showTextAligned(writer.getDirectContent(),
                    Element.ALIGN_CENTER, new Phrase("DEMO VERSION", FONT),
                    297.5f, 421, 45);
        }
    }

}
