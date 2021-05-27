package com.grapecity.spreadjsworld.Controller;

import com.grapecity.documents.excel.*;
import com.grapecity.documents.excel.drawing.HtmlSaveOptions;
import com.grapecity.documents.excel.drawing.ImageType;
import com.grapecity.spreadjsworld.Model.*;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.pdmodel.PDPage;
import org.apache.pdfbox.pdmodel.common.PDRectangle;
import org.apache.pdfbox.rendering.PDFRenderer;
import org.springframework.http.HttpHeaders;
import org.springframework.http.HttpStatus;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.util.ResourceUtils;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;

import javax.imageio.ImageIO;
import java.awt.image.BufferedImage;
import java.io.*;
import java.net.URL;
import java.util.ArrayList;
import java.util.Date;
import java.util.EnumSet;
import java.util.List;
import java.util.zip.GZIPInputStream;
import java.util.zip.GZIPOutputStream;

@RestController
@RequestMapping({"Spread", "spread"})
public class SpreadController {

    static {
        //Workbook.SetLicenseKey("");
    }

    public SpreadController() {
//        指定字体文件路径
//        Workbook.FontsFolderPath = "/Users/dexteryao/Documents/Projects/IdeaProjects/common/pdfFont";
    }

    /*获取reports路径下模板文件列表，返回分号分割文件名string*/
    @CrossOrigin
    @RequestMapping(value = "/getTemplates", method = RequestMethod.GET)
    public String getTemplates() {

        try {
            String path = "src/main/resources/reports";
            File reportDir = ResourceUtils.getFile(path);
            File[] files = reportDir.listFiles();
            StringBuilder result = new StringBuilder();
            for (File file : files) {
                if (file.isFile() && !file.isHidden()) {
                    result.append(file.getName()).append(";");
                        System.out.println("^^^^^" + file.getName());
                }
            }
            result = new StringBuilder(compress(result.toString()));
            return result.toString();
        } catch (Exception e) {
            e.printStackTrace();
        }
        return "0";
    }

    /*通过文件名，返回压缩后json，如果是Excel文件，使用GCExcel toJSON*/
    @CrossOrigin
    @RequestMapping(value = "/loadTemplate", method = RequestMethod.POST)
    public String loadTemplate(@RequestParam("fileName") String fileName) {
        try {
            String filePath = "src/main/resources/reports/" + fileName;
            String workbookJSON;
            Workbook workbook = new Workbook();

            System.out.println("开始获取数据：" + new Date());
            if (fileName.endsWith("xlsx")) {
                URL url = ResourceUtils.getURL(filePath);
                workbook.open(url.getFile());
                workbook.save("temp/temp.xlsx");
                workbookJSON = workbook.toJson();
                BufferedWriter out2 = new BufferedWriter(new FileWriter("temp/test.ssjson"));
                out2.write(workbookJSON);
                out2.close();



            } else {
                workbookJSON = getTextFileContent(filePath);
            }

            System.out.println("开始压缩：" + new Date());
            String result = compress(workbookJSON);

            System.out.println("压缩结束：" + new Date());
            return result;
        } catch (Exception e) {
            e.printStackTrace();
        }
        return "0";
    }

    /*上传文件，报错到reports文件夹，excel和ssjson都通过GCExcel 处理一遍后再返回压缩后ssjson，用于测试GCExcel */
    @CrossOrigin
    @RequestMapping(value = "/uploadFile", headers = ("content-type=multipart/form-data"), method = RequestMethod.POST)
    public String uploadFile(@RequestParam("file") MultipartFile file,
                             @RequestParam("fileName") String fileName) throws FileNotFoundException {
        if (file.isEmpty()) {
            System.out.println("文件空");
            return "上传失败！";
        }

        String filePath = "src/main/resources/reports/" + fileName;
        URL url = ResourceUtils.getURL(filePath);
        System.out.println(filePath);
        File dest = new File(url.getFile());
        try {
            file.transferTo(dest);

            Workbook workbook = new Workbook();

            if (fileName.endsWith("xlsx")) {
                workbook.open(url.getFile());
                workbook.save("temp/test.pdf", SaveFileFormat.Pdf);
            } else {
                String workbookJSON = getTextFileContent(url.getFile());
                workbook.fromJson(workbookJSON);
            }

            String workbookJSON = workbook.toJson();
            String result = compress(workbookJSON);

            return result;
        } catch (IOException e) {
            e.printStackTrace();
        }
        return "上传失败！";
    }

    /* 导出PDF，可选择单sheet导出或者绑定模拟数据后导出*/
    @CrossOrigin
    @RequestMapping(value = "/getpdf", method = RequestMethod.POST)
    public ResponseEntity<byte[]> getPDF(@RequestParam(value = "data", required = true) String data,
                                         @RequestParam(value = "bindingType", required = false) int bindingType,
                                         @RequestParam(value = "isActiveSheet", required = false) boolean isActiveSheet) {

        String json = uncompress(data);
        Workbook workbook = new Workbook();
        workbook.setEnableCalculation(true);
        workbook.fromJson(json);

        // type 1 GCExcel 模板绑定， 2 GCExcel 使用SpreadJS binding path 绑定
        if (bindingType == 1) {
            bindingDataToWorkbook(workbook);
        }
        else if(bindingType == 2){
            IWorksheet sheet = workbook.getActiveSheet();
            bindingSpreadJSDataToWorkSheet(sheet);
        }

        ByteArrayOutputStream out = new ByteArrayOutputStream();

        if (isActiveSheet) {
            IWorksheet sheet = workbook.getActiveSheet();
            adjustSheetForPDF(sheet);
            sheet.save(out, SaveFileFormat.Pdf);
        } else {
            workbook.getWorksheets().forEach((sheet)->{
                adjustSheetForPDF(sheet);
            });
            workbook.save(out, SaveFileFormat.Pdf);
        }
//        workbook.save("temp/test.xlsx");
//        System.out.print(workbook.toJson());

        byte[] contents = out.toByteArray();

        HttpHeaders headers = new HttpHeaders();
        headers.setContentType(MediaType.APPLICATION_PDF);
        // Here you have to set the actual filename of your pdf
        String filename = "output.pdf";
        headers.setContentDispositionFormData(filename, filename);
        headers.setCacheControl("must-revalidate, post-check=0, pre-check=0");
        ResponseEntity<byte[]> response = new ResponseEntity<>(contents, headers, HttpStatus.OK);
        return response;
    }

    private void getSheetImage(OutputStream out, IWorksheet sheet, boolean isSelectedRange, boolean isPDFMode) throws IOException {
        if(isPDFMode){
            if(isSelectedRange){
                sheet.getPageSetup().setPrintArea(sheet.getSelection().toString());
            }

            //Create a PrintManager.
            PrintManager printManager = new PrintManager();


            //Get the first print area of the worksheet.
            IRange printArea = printManager.getPrintAreas(sheet).get(0);

            //Get the size of the printArea.
            Size size = printManager.getSize(printArea);
            PDRectangle pageSize = new PDRectangle((float) size.getWidth() + 10, (float) size.getHeight() + 10);

            //Create a pdf document.
            PDDocument doc = new PDDocument();
            PDPage page = new PDPage(pageSize);
            doc.addPage(page);

            //Draw the print to the specified location on the page.
            printManager.draw(doc, page, new Point(5, 5), printArea);


            ByteArrayOutputStream tempStream = new ByteArrayOutputStream();
            doc.save(tempStream);
            tempStream.close();

            //Saves the page as an image to a stream.
            PDFRenderer pdfRenderer = new PDFRenderer(doc);
            BufferedImage bim;

            bim = pdfRenderer.renderImageWithDPI(0, 96*2);
            ImageIO.write(bim, "PNG", out);
            doc.close();

        }
        else{
            if(isSelectedRange){
                sheet.getSelection().toImage(out, ImageType.PNG);
            }
            else {
                sheet.toImage(out, ImageType.PNG);
            }
        }
    }

    /* 导出图片，可选择单sheet导出或者选择区域导出*/
    @CrossOrigin
    @RequestMapping(value = "/getImage", method = RequestMethod.POST)
    public ResponseEntity<byte[]> getImage(@RequestParam(value = "data", required = true) String data,
                                           @RequestParam(value = "isPDFMode", required = false)boolean isPDFMode,
                                           @RequestParam(value = "isSelectedRange", required = false) boolean isSelectedRange) {

        String json = uncompress(data);
        Workbook workbook = new Workbook();
        workbook.setEnableCalculation(true);
        workbook.fromJson(json);

        IWorksheet sheet = workbook.getActiveSheet();
        adjustSheetForPDF(sheet);
        ByteArrayOutputStream out = new ByteArrayOutputStream();

        try {
            getSheetImage(out, sheet, isSelectedRange, isPDFMode);
        }
        catch (Exception e){
            e.printStackTrace();
        }

        byte[] contents = out.toByteArray();
        HttpHeaders headers = new HttpHeaders();
        headers.setContentType(MediaType.IMAGE_PNG);
        // Here you have to set the actual filename of your pdf
        String filename = "image.png";
        headers.setContentDispositionFormData(filename, filename);
        headers.setCacheControl("must-revalidate, post-check=0, pre-check=0");
        ResponseEntity<byte[]> response = new ResponseEntity<>(contents, headers, HttpStatus.OK);
        return response;
    }

    /*绑定模拟数据*/
    @CrossOrigin
    @RequestMapping(value = "/bindingData", method = RequestMethod.POST)
    public String bindingData(@RequestParam(value = "data", required = true) String data) {
        try {
            String json = uncompress(data);
            if (json != null && !json.equals("")) {

                Workbook workbook = new Workbook();
                workbook.setEnableCalculation(false);
                workbook.fromJson(json);

                bindingDataToWorkbook(workbook);

                System.out.println("开始toJSON：" + new Date());
                String workbookJSON = workbook.toJson();
                System.out.println("开始压缩：" + new Date());
                String result = compress(workbookJSON);
                System.out.println("结束：" + new Date());
//                workbook.save(ResourceUtils.getURL("src/main/resources/reports/result.xlsx").getPath());
                return result;
            }
            return "0";
        } catch (Exception e) {
            // TODO Auto-generated catch block
            e.printStackTrace();
        }
        return "0";
    }

    @CrossOrigin
    @RequestMapping(value = "/getHTML", method = RequestMethod.POST)
    public ResponseEntity<byte[]> getHTML(@RequestParam(value = "data", required = true) String data) {
        try {
            String json = uncompress(data);
            if (json != null && !json.equals("")) {

                Workbook workbook = new Workbook();
                workbook.setEnableCalculation(false);
                workbook.fromJson(json);
                IWorksheet sheet = workbook.getActiveSheet();
                HtmlSaveOptions options = new HtmlSaveOptions();

                options.setExportSheetName(sheet.getName());
                // Set exported image as base64
                options.setExportImageAsBase64(true);
                // Set exported css style in html file
                options.setExportCssSeparately(false);
                // Set not to export single tab in html
                options.setExportSingleTab(false);

                ByteArrayOutputStream tempStream = new ByteArrayOutputStream();
                workbook.save(tempStream, options);

                byte[] contents = tempStream.toByteArray();
                HttpHeaders headers = new HttpHeaders();
                headers.setContentType(MediaType.TEXT_HTML);
                // Here you have to set the actual filename of your pdf
                String filename = "index.html";
                headers.setContentDispositionFormData(filename, filename);
                headers.setCacheControl("must-revalidate, post-check=0, pre-check=0");
                ResponseEntity<byte[]> response = new ResponseEntity<>(contents, headers, HttpStatus.OK);
                return response;
            }
            return null;
        } catch (Exception e) {
            // TODO Auto-generated catch block
            e.printStackTrace();
        }
        return null;
    }


    @CrossOrigin
    @RequestMapping(value = "/bindingSJSData", method = RequestMethod.POST)
    public String bindingSJSData(@RequestParam(value = "data", required = true) String data) {
        try {
            String json = uncompress(data);
            if (json != null && !json.equals("")) {

                Workbook workbook = new Workbook();
                workbook.setEnableCalculation(false);
                workbook.fromJson(json);

                IWorksheet sheet = workbook.getActiveSheet();
                bindingSpreadJSDataToWorkSheet(sheet);

                System.out.println("开始toJSON：" + new Date());
                String workbookJSON = workbook.toJson();
                System.out.println("开始压缩：" + new Date());
                String result = compress(workbookJSON);
                System.out.println("结束：" + new Date());
//                workbook.save(ResourceUtils.getURL("src/main/resources/reports/result.xlsx").getPath());
                return result;
            }
            return "0";
        } catch (Exception e) {
            // TODO Auto-generated catch block
            e.printStackTrace();
        }
        return "0";
    }

    public void bindingDataToWorkbook(Workbook workbook) {

        System.out.println("开始创建模拟数据：" + new Date());
        List<DataSourceModel> ds1 = new ArrayList<>();
        for (int i = 0; i < 30; i++) {
            for (int j = 0; j < 77; j++) {
                for (int k = 0; k < 2; k++) {
                    DataSourceModel dsm = new DataSourceModel();
                    dsm.Area = "地区" + j;
                    dsm.BaseAmount = i * j * 6 + k;
                    dsm.InsuranceYear = 2020;
                    dsm.SecondaryInstitution = "机构" + j + k;
                    dsm.InsuranceCount = (j % 5) * i + k;
                    dsm.EndMonth = "2020-" + (k % 2 + 8);
                    dsm.EndDate = dsm.EndMonth + "-" + (i + 1);
                    ds1.add(dsm);
                }
            }
        }

        List<CommonItemModel> incomeDetail = new ArrayList<>();
        for (int i = 0; i < 3; i++) {
            CommonItemModel item = new CommonItemModel();
            item.ItemName = "收入来源" + i;
            item.ItemValue = (int)(Math.random()*(10000));
            incomeDetail.add(item);
        }

        List<CommonItemModel> ExpensesDetail = new ArrayList<>();
        for (int i = 0; i < 5; i++) {
            CommonItemModel item = new CommonItemModel();
            item.ItemName = "每月支出" + i;
            item.ItemValue = (int)(Math.random()*(1000));
            ExpensesDetail.add(item);
        }
        List<CommonItemModel> SavingsDetail = new ArrayList<>();
        for (int i = 0; i < 5; i++) {
            CommonItemModel item = new CommonItemModel();
            item.ItemName = "每月存款" + i;
            item.ItemValue = (int)(Math.random()*(5000));
            SavingsDetail.add(item);
        }

        System.out.println("模拟数据行数：" + ds1.size());
        System.out.println("开始添加数据：" + new Date());
        workbook.addDataSource("ds", ds1);
        workbook.addDataSource("Income", incomeDetail);
        workbook.addDataSource("Expenses", ExpensesDetail);
        workbook.addDataSource("Savings", SavingsDetail);
        System.out.println("开始绑定数据：" + new Date());
        workbook.processTemplate();
        System.out.println("结束绑定数据：" + new Date());
    }

    public void bindingSpreadJSDataToWorkSheet(IWorksheet sheet){
        Company company1 = new Company("GrapeCity", null, "We know everything!", "Gaoxin 6th road", "Xi'an", "029-88331988", "grapecity@grapecity.com"),
                company2 = new Company("Tecent", null, "We have everything!", "Shenzhen 2st road", "Shenzhen", "0755-12345678", "tecent@qq.com"),
                company3 = new Company("Alibaba", null, "We sale everything!", "Hangzhou 3rd road", "Hangzhou", "0571-12345678", "alibaba@alibaba.com");
        Customer customer1 = new Customer("A1", "employee 1", company2),
                customer2 = new Customer("A2", "employee 2", company3);
        List<Record> records1 = new ArrayList<>();
        Invoice invoice1 = new Invoice(company1, "00001", new Date(2014, 0, 1), customer1, customer2, records1);

        for (int i = 0; i < Math.random() * 20; i++) {
            records1.add(new Record("Description" + i, i, Math.round(Math.random() * 150)));
        }
        sheet.setDataSource(invoice1);
    }


    public static void adjustSheetForPDF(IWorksheet sheet){
        PrintManager printManager = new PrintManager();
        IRange printArea = printManager.getPrintAreas(sheet).get(0);
        double rangeWidth = 0, rangeHeight = 0;
        for(int row = printArea.getRow(); row < printArea.getRow() + printArea.getRowCount(); row++){
            rangeHeight += sheet.getRows().get(row).getRowHeightInPixel();
        }
        for(int col = printArea.getColumn(); col < printArea.getColumn() + printArea.getColumnCount(); col++){
            rangeWidth += sheet.getColumns().get(col).getColumnWidthInPixel();
        }
        System.out.println("Range:" + rangeHeight + "," + rangeWidth);
        System.out.println("Range:" + printArea);
        System.out.println("Range:" + rangeHeight + "," + rangeWidth);
        Size size = printManager.getSize(printArea);

        System.out.println("Size:" + size.getHeight() + "," + size.getWidth());
        double rate = size.getHeight() * rangeWidth / (rangeHeight * 0.96 * size.getWidth());
        System.out.println("Rate:" + rate);
        for(int col = 0; col < printArea.getColumnCount(); col++){
            sheet.getColumns().get(col).setColumnWidthInPixel(Math.round(sheet.getColumns().get(col).getColumnWidthInPixel() * rate + 0.4));
        }
        for(int row = 0; row < printArea.getRowCount() - printArea.getRow(); row++){
            printArea.getRows().get(row).setRowHeightInPixel(printArea.getRows().get(row).getRowHeightInPixel() / 0.96);
        }
    }
    public static String compress(String str) {
        if (str.length() <= 0) {
            return str;
        }
        try {
            ByteArrayOutputStream bos = null;
            GZIPOutputStream os = null; //使用默认缓冲区大小创建新的输出流
            byte[] bs = null;
            try {
                bos = new ByteArrayOutputStream();
                os = new GZIPOutputStream(bos);
                os.write(str.getBytes()); //写入输出流
                os.close();
                bos.close();
                bs = bos.toByteArray();
                return new String(bs, "ISO-8859-1");
            } finally {
                bs = null;
                bos = null;
                os = null;
            }
        } catch (Exception ex) {
            return str;
        }
    }

    public static String uncompress(String str) {
        if (str.length() <= 0) {
            return str;
        }
        try {
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            ByteArrayInputStream in = new ByteArrayInputStream(str.getBytes("ISO-8859-1"));
            GZIPInputStream ungzip = new GZIPInputStream(in);
            byte[] buffer = new byte[256];
            int n;
            while ((n = ungzip.read(buffer)) >= 0) {
                out.write(buffer, 0, n);
            }
            return new String(out.toByteArray(), "UTF-8");
        } catch (Exception e) {

        }
        return str;
    }

    private String getTextFileContent(String path) {
        try {
            File file = ResourceUtils.getFile(path);
            String encoding = "UTF-8";
            Long fileLength = file.length();
            byte[] fileContent = new byte[fileLength.intValue()];
            FileInputStream in = new FileInputStream(file);
            in.read(fileContent);
            in.close();
            String content = new String(fileContent, encoding);
            return content;
        } catch (Exception e) {

        }
        return "";
    }

}
