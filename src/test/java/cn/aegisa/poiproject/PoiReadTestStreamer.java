package cn.aegisa.poiproject;

import lombok.extern.slf4j.Slf4j;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.openxml4j.opc.PackageAccess;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.util.CellAddress;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.util.SAXHelper;
import org.apache.poi.xssf.eventusermodel.ReadOnlySharedStringsTable;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.eventusermodel.XSSFSheetXMLHandler;
import org.apache.poi.xssf.model.StylesTable;
import org.apache.poi.xssf.usermodel.XSSFComment;
import org.junit.Test;
import org.junit.runner.RunWith;
import org.springframework.boot.test.context.SpringBootTest;
import org.springframework.test.context.junit4.SpringRunner;
import org.xml.sax.ContentHandler;
import org.xml.sax.InputSource;
import org.xml.sax.XMLReader;

import java.io.File;
import java.io.InputStream;
import java.io.PrintStream;

@RunWith(SpringRunner.class)
@SpringBootTest
@Slf4j
public class PoiReadTestStreamer {

    @Test
    public void contextLoads() throws Exception {
        final File file = new File("D:\\hnair.xlsx");
        final OPCPackage aPackage = OPCPackage.open(file, PackageAccess.READ);
        final XSSFReader reader = new XSSFReader(aPackage);
        // 这两个参数不知道是做什么的
        StylesTable styles = reader.getStylesTable();
        ReadOnlySharedStringsTable strings = new ReadOnlySharedStringsTable(aPackage);
        final XSSFReader.SheetIterator iterator = (XSSFReader.SheetIterator) reader.getSheetsData();
        while (iterator.hasNext()) {
            final InputStream inputStream = iterator.next();
            final String sheetName = iterator.getSheetName();
            log.info("获取到一个sheet：{}", sheetName);
            InputSource sheetSource = new InputSource(inputStream);
            XMLReader sheetParser = SAXHelper.newXMLReader();
            ContentHandler handler = new XSSFSheetXMLHandler(
                    styles, null, strings, new SheetHandler(), new DataFormatter(), false);

            sheetParser.setContentHandler(handler);
            sheetParser.parse(sheetSource);

            inputStream.close();
        }
    }

    private int minColumns = 0;
    private PrintStream output = System.out;

    private class SheetHandler implements XSSFSheetXMLHandler.SheetContentsHandler {
        private boolean firstCellOfRow = false;
        private int currentRow = -1;
        private int currentCol = -1;

        private void outputMissingRows(int number) {
            for (int i = 0; i < number; i++) {
                for (int j = 0; j < minColumns; j++) {
                    output.append(',');
                }
                output.append('\n');
            }
        }

        @Override
        public void startRow(int rowNum) {
            // If there were gaps, output the missing rows
            outputMissingRows(rowNum - currentRow - 1);
            // Prepare for this row
            firstCellOfRow = true;
            currentRow = rowNum;
            currentCol = -1;
        }

        @Override
        public void endRow(int rowNum) {
            // Ensure the minimum number of columns
            for (int i = currentCol; i < minColumns; i++) {
                output.append(',');
            }
            output.append('\n');
        }

        @Override
        public void cell(String cellReference, String formattedValue,
                         XSSFComment comment) {
            if (firstCellOfRow) {
                firstCellOfRow = false;
            } else {
                output.append(',');
            }

            // gracefully handle missing CellRef here in a similar way as XSSFCell does
            if (cellReference == null) {
                cellReference = new CellAddress(currentRow, currentCol).formatAsString();
            }

            // Did we miss any cells?
            int thisCol = (new CellReference(cellReference)).getCol();
            int missedCols = thisCol - currentCol - 1;
            for (int i = 0; i < missedCols; i++) {
                output.append(',');
            }
            currentCol = thisCol;

            // Number or string?
            try {
                Double.parseDouble(formattedValue);
                output.append(formattedValue);
            } catch (NumberFormatException e) {
                output.append('"');
                output.append(formattedValue);
                output.append('"');
            }
        }

        @Override
        public void headerFooter(String text, boolean isHeader, String tagName) {
            // Skip, no headers or footers in CSV
        }
    }


}
