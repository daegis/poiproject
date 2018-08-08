package cn.aegisa.poiproject;

import lombok.extern.slf4j.Slf4j;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Test;
import org.junit.runner.RunWith;
import org.springframework.boot.test.context.SpringBootTest;
import org.springframework.core.io.ClassPathResource;
import org.springframework.test.context.junit4.SpringRunner;

import java.io.File;
import java.io.FileOutputStream;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.UUID;

@RunWith(SpringRunner.class)
@SpringBootTest
@Slf4j
public class PoiprojectApplicationTests {

    @Test
    public void contextLoads() throws Exception {
        ClassPathResource resource = new ClassPathResource("excel/hnaorder.xlsx");
        XSSFWorkbook workbook = new XSSFWorkbook(resource.getInputStream());
        XSSFSheet sheet = workbook.getSheetAt(0);
        log.info("初始化海航订单模板成功");
        int rowIndex = 1;
        int startOrderNo = 1000000;
        for (int i = 0; i < 1000000; i++) {
            if (i % 10000 == 0) {
                System.out.println(i);
            }
            XSSFRow row = sheet.createRow(rowIndex++);
            row.createCell(0).setCellValue("HNA" + startOrderNo++);
            row.createCell(1).setCellValue("机票");
            row.createCell(2).setCellValue(Math.random() * 555);
            row.createCell(3).setCellValue(Math.random() * 555);
            LocalDateTime orderTime = LocalDateTime.now();
            row.createCell(4).setCellValue(orderTime.format(DateTimeFormatter.ofPattern("yyyy-MM-dd HH:mm:ss")));
            row.createCell(5).setCellValue("折扣机票");
            row.createCell(6).setCellValue(UUID.randomUUID().toString());
            row.createCell(7).setCellValue(Math.random() * 555);
            row.createCell(8).setCellValue("海航");
        }

        try (FileOutputStream outputStream = new FileOutputStream(new File("d:\\hnair.xlsx"))) {
            workbook.write(outputStream);
        } catch (Exception e) {
            throw new RuntimeException(e);
        }
    }

}
