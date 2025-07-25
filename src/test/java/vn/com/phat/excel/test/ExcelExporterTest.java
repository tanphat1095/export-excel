/*
 * Copyright 2025 tanphat.1095
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *     http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */
package vn.com.phat.excel.test;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.CellStyle;
import org.junit.Assert;
import org.junit.Test;
import vn.com.phat.excel.ExcelExporter;
import vn.com.phat.excel.ExcelExporterDefaultImpl;
import vn.com.phat.excel.cellhandler.CellDataTypeHandler;
import vn.com.phat.excel.cellstyle.CellStyleHandler;
import vn.com.phat.excel.dto.ItemColsExcelDto;
import vn.com.phat.excel.enumdef.JavaDataType;
import vn.com.phat.excel.param.ExcelExporterParam;
import vn.com.phat.excel.param.ExcelExporterParamDefault;
import vn.com.phat.excel.param.ExcelExporterSheetParam;
import vn.com.phat.excel.param.ExcelExporterSheetParamDefault;
import vn.com.phat.excel.util.ExcelColumnExtractor;
import vn.com.phat.excel.util.ExcelDataExtractor;

import java.io.File;
import java.io.FileOutputStream;
import java.lang.reflect.Field;
import java.math.BigDecimal;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Objects;
import java.util.UUID;

/**
 * Default implementation of the ExcelExporter interface.
 * This class provides methods to export data to an Excel file using Apache POI library.
 * It supports different data types and cell styles.
 *
 * @author phatlt
 */
public class ExcelExporterTest {

    // Write data to excel with a large data set, and multiple sheet
    // Additional action after the writing process has done.
    @Test
    public void test() throws Exception {
        File file = new File(getClass().getResource(".").getPath() + "/test.xlsx");
        try(FileOutputStream fos = new FileOutputStream(file)) {
            ExcelExporter exporter = getExcelExporter();

            List<ItemColsExcelDto> cols = new ArrayList<>();
            ExcelColumnExtractor.extractColumn(ExcelTestEnum.class, cols);

            List<ExcelTestData> datas = new ArrayList<>();
            for( int i = 0; i < 100000; i ++){
                ExcelTestData data = new ExcelTestData(i+1, UUID.randomUUID().toString(), new Date(), i % 2 == 0, LocalDate.now(), new BigDecimal(
                        "1.02"), UUID.randomUUID().toString());
                datas.add(data);
            }

            long startExtract = System.currentTimeMillis();
            List<Map<String,Object>> mappingData = ExcelDataExtractor.extractData(datas, ExcelTestData.class);
            System.out.println(System.currentTimeMillis() - startExtract);
            Map<String, Field> fieldMapping = ExcelDataExtractor.extractFieldMapping(ExcelTestData.class);

            ExcelExporterSheetParam sheet = ExcelExporterSheetParamDefault.builder()
                    .startCell("A1")
                    .colIndex(cols)
                    .data(mappingData)
                    .sheetIndex(0)
                    .mapFields(fieldMapping)
                    .build();

            ExcelExporterSheetParam sheet2 = ExcelExporterSheetParamDefault.builder()
                    .startCell("A5")
                    .colIndex(cols)
                    .data(mappingData)
                    .sheetIndex(1)
                    .mapFields(fieldMapping)
                    .sheetName("SecondSheet")
                    .build();

            ExcelExporterSheetParam sheet3 = ExcelExporterSheetParamDefault.builder()
                    .startCell("A10")
                    .colIndex(cols)
                    .data(mappingData)
                    .sheetIndex(3)
                    .mapFields(fieldMapping)
                    .sheetName("Third")
                    .build();

            ExcelExporterSheetParam sheet4 = ExcelExporterSheetParamDefault.builder()
                    .startCell("A7")
                    .colIndex(cols)
                    .data(mappingData)
                    .sheetIndex(4)
                    .mapFields(fieldMapping)
                    .sheetName("Fourth")
                    .build();

            List<ExcelExporterSheetParam> sheets = Arrays.asList(sheet, sheet2, sheet3, sheet4);
            ExcelExporterParam excelParam = ExcelExporterParamDefault.builder()
                    .consumer(workbook-> workbook.setSheetName(0, "FirstSheet"))
                    .templatePath(Objects.requireNonNull(this.getClass().getClassLoader().getResource("template.xlsx")).getPath())
                    .sheets(sheets)
                    .build();
            long start = System.currentTimeMillis();
            exporter.doExport(fos, excelParam);
            System.out.println(System.currentTimeMillis() - start);
            fos.flush();
            Assert.assertNotNull(fos);
        }
        // One million rows data and 2 sheets will take around 30 seconds for a writing process
        // 100k rows data and 4 sheets will take less than 5 seconds
        // I've just added the feature to support a multiple sheet in multi-thread processing.
        // This reduces the time to write data to excel.
    }

    // Add data type handler doesn't exist in the default implementation.
    // We can change the data type handler's behavior by overriding its default handler
    private static ExcelExporter getExcelExporter() {
        CellDataTypeHandler localDateHandler = (cell, value) -> {
            DateTimeFormatter formatter = DateTimeFormatter.ofPattern("dd/MM/yyyy");
            String cellValue = ((LocalDate) value).format(formatter);
            cell.setCellValue(cellValue);
        };
        Map<String, CellDataTypeHandler> cellDataHandler = new HashMap<>();
        cellDataHandler.put("LOCALDATE",localDateHandler);

        // You can override the handler already exist.
        CellDataTypeHandler stringHandler = (cell, value) ->{
          String cellValue = (String) value + "_STRING";
          cell.setCellValue(cellValue);
        };

        // You can also do the same thing with field name.
        CellDataTypeHandler stringFieldHandler = (cell, value) ->{
            String cellValue = (String) value + "_StringField";
            cell.setCellValue(cellValue);
        };

        Map<String, CellStyleHandler> cellStyleCustom = getStringCellStyleHandlerMap();

        cellDataHandler.put(JavaDataType.STRING.name(), stringHandler);
        cellDataHandler.put("CUSTOMSTYLE", stringFieldHandler);

        return new ExcelExporterDefaultImpl(cellDataHandler, cellStyleCustom);
    }

    private static Map<String, CellStyleHandler> getStringCellStyleHandlerMap() {
        CellStyleHandler customStyleHandler = (workbook, dataFormat) -> {
            CellStyle cellStyle = workbook.createCellStyle();
            cellStyle.setAlignment(org.apache.poi.ss.usermodel.HorizontalAlignment.CENTER);
            cellStyle.setBorderTop(BorderStyle.DASH_DOT);
            cellStyle.setBorderBottom(BorderStyle.DASH_DOT);
            cellStyle.setBorderLeft(BorderStyle.DASH_DOT);
            cellStyle.setBorderRight(BorderStyle.DASH_DOT);
            return cellStyle;
        };
        Map<String, CellStyleHandler> cellStyleCustom = new HashMap<>();
        cellStyleCustom.put("CUSTOMSTYLE", customStyleHandler);
        return cellStyleCustom;
    }
}
