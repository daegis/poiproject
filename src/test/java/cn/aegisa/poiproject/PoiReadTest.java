package cn.aegisa.poiproject;

import lombok.extern.slf4j.Slf4j;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Test;
import org.junit.runner.RunWith;
import org.springframework.boot.test.context.SpringBootTest;
import org.springframework.test.context.junit4.SpringRunner;

import java.io.File;
import java.io.FileInputStream;

@RunWith(SpringRunner.class)
@SpringBootTest
@Slf4j
public class PoiReadTest {

    @Test
    public void contextLoads() throws Exception {
        final FileInputStream fileInputStream = new FileInputStream(new File("D:\\hnair.xlsx"));
        XSSFWorkbook workbook = new XSSFWorkbook(fileInputStream);
        XSSFSheet sheet = workbook.getSheetAt(0);
        final int number = sheet.getPhysicalNumberOfRows();
        System.out.println(number);
        Thread.sleep(99999L);
    }

}
