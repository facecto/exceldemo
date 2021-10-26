package com.facecto.exceldemo.controller;

import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.ExcelWriter;
import com.alibaba.excel.write.metadata.WriteSheet;
import com.alibaba.excel.write.metadata.fill.FillConfig;
import com.facecto.code.base.CodeException;
import com.facecto.exceldemo.entity.*;
import com.github.liaochong.myexcel.core.ExcelBuilder;
import com.github.liaochong.myexcel.core.FreemarkerExcelBuilder;
import com.github.liaochong.myexcel.utils.AttachmentExportUtil;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.core.io.ClassPathResource;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

import javax.servlet.http.HttpServletResponse;
import java.io.*;
import java.math.BigDecimal;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.zip.ZipEntry;
import java.util.zip.ZipOutputStream;

/**
 * @author Jon So, https://cto.pub, https://github.com/facecto
 * @version v1.0.0 (2021/10/23)
 */
@RestController
public class ReportController {
    @GetMapping("/01")
    public void rep01ex(HttpServletResponse response) {
        Map<String, Object> dataMap = new HashMap<>();
        dataMap.put("sheetName", "example01");
        dataMap.put("title", "example01");

        List<String> titles = new ArrayList<>();
        titles.add("id");
        titles.add("customer");
        titles.add("orderNo.");
        titles.add("orderType");
        titles.add("space");

        titles.add("pay");
        titles.add("date1");
        titles.add("money");
        titles.add("date2");
        titles.add("date3");

        titles.add("money2");
        titles.add("money3");
        titles.add("status1");
        titles.add("expire");
        titles.add("profit");

        titles.add("profit 1");
        titles.add("profit 2");
        titles.add("profit 3");
        titles.add("profit 4");
        titles.add("profit 5");

        titles.add("audit");

        dataMap.put("titles", titles);
        dataMap.put("tableMan","Made by Jon");
        dataMap.put("tableTime", "Create Date："+ LocalDate.now());
        dataMap.put("data",getReport1Data(1, "2021-10"));
        try(ExcelBuilder excelBuilder = new FreemarkerExcelBuilder()){
            Workbook workbook = excelBuilder.template("/templates/report01.ftl").build(dataMap);
            AttachmentExportUtil.export(workbook, "report01", response);
        }catch (Exception e){
            throw new CodeException("Error!");
        }
    }

    @GetMapping("/02")
    public void rep02ex(HttpServletResponse response) {
        Map<String, Object> dataMap = new HashMap<>();
        dataMap.put("sheetName", "example02");
        dataMap.put("title","example02");

        List<String> titles = new ArrayList<>();
        titles.add("ID");
        titles.add("Customer");
        titles.add("OrderNO.");
        titles.add("Type");
        titles.add("Space");
        titles.add("Pay");
        titles.add("Date1");

        titles.add("Supplier");
        titles.add("Bank");
        titles.add("Account");
        titles.add("AccountNO.");
        titles.add("Fee1");
        titles.add("Time");

        dataMap.put("titles", titles);
        dataMap.put("tableMan","Made by Jon");
        dataMap.put("tableTime", "Create Time："+ LocalDate.now());
        dataMap.put("data",getReport2Data(1, "2010-10"));
        try(ExcelBuilder excelBuilder = new FreemarkerExcelBuilder()){
            Workbook workbook = excelBuilder.template("/templates/report02.ftl").build(dataMap);
            AttachmentExportUtil.export(workbook, "report02", response);
        }catch (Exception ex){
            throw new CodeException("Error!");
        }
    }

    @GetMapping("/03")
    public void rep03ex(HttpServletResponse response) {
        Map<String, Object> dataMap = new HashMap<>();
        dataMap.put("sheetName", "example04");
        dataMap.put("title", "report");
        dataMap.put("tableMan","Made By Jon");
        dataMap.put("tableTime", "Create time："+ LocalDate.now());
        dataMap.put("data", getReport3Data(1, "2010-10"));

        try(ExcelBuilder excelBuilder = new FreemarkerExcelBuilder()){
            Workbook workbook = excelBuilder.template("/templates/report03.ftl").build(dataMap);
            AttachmentExportUtil.export(workbook, "report03", response);
        }catch (Exception ex){
            throw new CodeException("Error!");
        }
    }

    @GetMapping("/04")
    public void rep04ex(HttpServletResponse response) throws IOException {
        response.setHeader("Content-disposition", "attachment; filename=" + getName()+".xlsx");
        response.setHeader("Pragma", "No-cache");
        response.setHeader("Cache-Control", "no-cache");
        response.setDateHeader("Expires", 0);
        response.setContentType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;charset=UTF-8");

        ClassPathResource classPathResource = new ClassPathResource("templates/report4.xlsx");
        XSSFWorkbook xssfWorkbook = new XSSFWorkbook(classPathResource.getInputStream());

        copyFirstSheet(xssfWorkbook, 9);
        for (int i = 0; i < xssfWorkbook.getNumberOfSheets(); i++) {
            String sheetName = "file-" + (i + 1);
            xssfWorkbook.setSheetName(i, sheetName);
        }
        ByteArrayOutputStream outStream = new ByteArrayOutputStream();
        xssfWorkbook.write(outStream);
        ByteArrayInputStream inputStream = new ByteArrayInputStream(outStream.toByteArray());
        ExcelWriter excelWriter = EasyExcel.write(response.getOutputStream()).withTemplate(inputStream).build();
        for (int i = 0; i < xssfWorkbook.getNumberOfSheets(); i++) {
            fillReport4Data(excelWriter, i);
        }
        excelWriter.finish();
    }

    @GetMapping("/05")
    public String rep05ex(HttpServletResponse response) throws IOException {
        response.setHeader("Content-disposition", "attachment; filename=" + getName()+".zip");
        response.setHeader("Pragma", "No-cache");
        response.setHeader("Cache-Control", "no-cache");
        response.setDateHeader("Expires", 0);
        String filePath = getPath() +  getName() + ".xlsx";
//        System.out.println(filePath);
//        String templateFilePath = getPath() + "templates"+ File.separator +"report4.xlsx";
//        XSSFWorkbook xssfWorkbook = new XSSFWorkbook(new FileInputStream(templateFilePath));

        ClassPathResource classPathResource = new ClassPathResource("templates/report4.xlsx");
        XSSFWorkbook xssfWorkbook = new XSSFWorkbook(classPathResource.getInputStream());

        copyFirstSheet(xssfWorkbook, 9);

        for (int i = 0; i < xssfWorkbook.getNumberOfSheets(); i++) {
            String sheetName = "File-" + (i + 1);
            xssfWorkbook.setSheetName(i, sheetName);
        }

        ByteArrayOutputStream outStream = new ByteArrayOutputStream();
        xssfWorkbook.write(outStream);

        ByteArrayInputStream inputStream = new ByteArrayInputStream(outStream.toByteArray());
        ExcelWriter excelWriter = EasyExcel.write(filePath).withTemplate(inputStream).build();
        for (int i = 0; i < xssfWorkbook.getNumberOfSheets(); i++) {
            fillReport4Data(excelWriter, i);
        }
        excelWriter.finish();

        OutputStream out = response.getOutputStream();
        ZipOutputStream zos = new ZipOutputStream(out);
        XSSFWorkbook workbook = new XSSFWorkbook(new FileInputStream(filePath));
        zos.putNextEntry(new ZipEntry(getName() + ".xlsx" ));
        ByteArrayOutputStream bos = new ByteArrayOutputStream();
        workbook.write(bos);
        bos.writeTo(zos);
        bos.close();
        zos.closeEntry();
        zos.close();
        return "ok";
    }

    private List<Report1> getReport1Data(Integer areaId, String yearMonth){
        try{
            DateTimeFormatter df = DateTimeFormatter.ofPattern("yyyy-MM-dd HH:mm:ss");
            LocalDateTime beginDate = LocalDateTime.parse(yearMonth + "-01 00:00:00",df);
            LocalDateTime endDate = beginDate.plusMonths(1).minusSeconds(1);
            List<Report1> list = new ArrayList<>();
            for (int i = 0; i < 10; i++) {
                Report1 r = new Report1()
                        .setCustomer("Smith"+ i)
                        .setAudit("Y")
                        .setStatus1("status")
                        .setSpace("China")
                        .setMoney("0.123")
                        .setProfit(BigDecimal.valueOf(i))
                        .setProfit5(BigDecimal.valueOf(10+i))
                        .setProfit4(BigDecimal.valueOf(20+i))
                        .setProfit3(BigDecimal.valueOf(30+i))
                        .setProfit2(BigDecimal.valueOf(40+i))
                        .setProfit1(BigDecimal.valueOf(50+i))
                        .setPay("paypal")
                        .setOrderType("typeA")
                        .setOrderNo("2010-001-001-"+i)
                        .setMoney3(BigDecimal.valueOf(1+i))
                        .setMoney2(BigDecimal.valueOf(2+i))
                        .setId(1+i)
                        .setExpire(0)
                        .setDate3("2010-01-01")
                        .setDate2("2010-10-11")
                        .setDate1("2020-10-15");
                list.add(r);
            }
            return list;
        }catch (Exception e){
            throw new CodeException("Error!");
        }

    }

    private List<Report2> getReport2Data(Integer areaId, String yearMonth){
        try{
            DateTimeFormatter df = DateTimeFormatter.ofPattern("yyyy-MM-dd HH:mm:ss");
            LocalDateTime beginDate = LocalDateTime.parse(yearMonth + "-01 00:00:00",df);
            LocalDateTime endDate = beginDate.plusMonths(1).minusSeconds(1);
            List<Report2> list = new ArrayList<>();
            for (int i = 0; i < 10; i++) {
                Report2 r = new Report2()
                        .setType("Type")
                        .setTime("2020-10-20")
                        .setSupplier("Jack"+i)
                        .setSpace("NewYork")
                        .setPay("Alipay")
                        .setOrderNO("No.123456"+i)
                        .setId(i)
                        .setFee1("123")
                        .setDate1("2021-11-11")
                        .setBank("Bank Of Coffee")
                        .setAccountNO("123-456-58-"+i)
                        .setAccount("Jack"+i)
                        .setCustomer("Bod");
                        ;
                list.add(r);
            }
            return list;
        } catch (Exception e){
            throw new CodeException("Error!");
        }
    }

    private List<Report3> getReport3Data(Integer areaId, String yearMonth){
        try{
            DateTimeFormatter df = DateTimeFormatter.ofPattern("yyyy-MM-dd HH:mm:ss");
            LocalDateTime beginDate = LocalDateTime.parse(yearMonth + "-01 00:00:00",df);
            LocalDateTime endDate = beginDate.plusMonths(1).minusSeconds(1);
            List<Report3> list = new ArrayList<>();
            for (int i = 0; i < 10; i++) {
                Report3 r = new Report3()
                        .setId(i)
                        .setOrderNo("2010-10-20-"+i)
                        .setCustomer("Jack")
                        .setDate1("2020-10-21")
                        .setOrderType("Book")
                        .setOrderMoney(BigDecimal.valueOf(100))
                        .setProfit0(BigDecimal.valueOf(i+10))
                        .setProfit1(BigDecimal.valueOf(i+10))
                        .setProfit2(BigDecimal.valueOf(i+10))
                        .setProfit3(BigDecimal.valueOf(i+10))
                        .setProfit4(BigDecimal.valueOf(i+10))
                        .setProfit5(BigDecimal.valueOf(i+10))
                        .setProfit6(BigDecimal.valueOf(i+10))
                        .setProfit7(BigDecimal.valueOf(i+10))
                        .setProfit8(BigDecimal.valueOf(i+10))
                        .setProfit9(BigDecimal.valueOf(i+10))
                        .setProfit10(BigDecimal.valueOf(i+10))
                        .setProfit11(BigDecimal.valueOf(i+10))
                        .setName1("Tom")
                        .setName2("Smith")
                        .setName3("White")
                        .setName5("Bob")
                        .setName4("Adam")
                        .setName6("Ford")
                ;
                list.add(r);
            }
            return list;
        } catch (Exception e){
            throw new CodeException("Error!");
        }
    }

    private String getName(){
        DateTimeFormatter formatter = DateTimeFormatter.ofPattern("yyyyMMdd-HHmmss");
        return formatter.format(LocalDateTime.now());
    }

    private static String getPath() {
        return RestController.class.getResource("/").getPath();
    }

    private void copyFirstSheet(XSSFWorkbook workbook, int times) {
        if (times <= 0)  throw new IllegalArgumentException("times error");
        for (int i = 0; i < times; i++) {
            workbook.cloneSheet(0);
        }
    }

    private void fillReport4Data(ExcelWriter excelWriter, int sheetNo) {
        Map<String, Object> map = new HashMap<>();
        map.put("title","RhinosLab.com");
        map.put("date1", "2021-10-22");
        map.put("author", "Jon");
        map.put("company", "RhinosLab -" + sheetNo);
        map.put("contact", "Jon");
        map.put("phone", "(001)123-456-789");
        map.put("email", "jon@rhinoslab.com");
        map.put("money", "12345.67");
        map.put("tax", "1123.46");
        map.put("percent", "1%");
        map.put("sign","Jon");
        map.put("date3","2021-10-23");
        WriteSheet writeSheet = EasyExcel.writerSheet(sheetNo).build();
        excelWriter.fill(map, writeSheet);

        FillConfig fillConfig = FillConfig.builder().forceNewRow(Boolean.TRUE).build();
        List<Report4> list = getReport4Data();
        excelWriter.fill(list, fillConfig, writeSheet);
    }

    private List<Report4> getReport4Data(){
        List<Report4> list = new ArrayList<>();
        for(int i=1;i<21;i++){
            Report4 r = new Report4()
                    .setNo(i)
                    .setGoods("Coke "+i)
                    .setNum(i)
                    .setPrice(BigDecimal.valueOf(2.5+i))
                    .setTax(BigDecimal.valueOf(0.01))
                    .setBarCode("2002-201-454-"+i)
                    .setPd("2021-10-21")
                    .setExp("2022-10-20")
                    .setSt(BigDecimal.valueOf(i).multiply(BigDecimal.valueOf(2.5+i)));
            list.add(r);
        }
        return list;
    }

}
