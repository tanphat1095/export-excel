package vn.com.phat.excel;

import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.streaming.SXSSFCell;
import org.apache.poi.xssf.streaming.SXSSFRow;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.util.Assert;
import vn.com.phat.excel.cellhandler.CellBigDecimalHandler;
import vn.com.phat.excel.cellhandler.CellBooleanHandler;
import vn.com.phat.excel.cellhandler.CellDataTypeHandler;
import vn.com.phat.excel.cellhandler.CellDateHandler;
import vn.com.phat.excel.cellhandler.CellDoubleHandler;
import vn.com.phat.excel.cellhandler.CellIntegerHandler;
import vn.com.phat.excel.cellhandler.CellLongHandler;
import vn.com.phat.excel.cellhandler.CellStringHandler;
import vn.com.phat.excel.cellstyle.CellStyleBigDecimalHandler;
import vn.com.phat.excel.cellstyle.CellStyleBooleanHandler;
import vn.com.phat.excel.cellstyle.CellStyleDateHandler;
import vn.com.phat.excel.cellstyle.CellStyleHandler;
import vn.com.phat.excel.cellstyle.CellStyleStringHandler;
import vn.com.phat.excel.dto.ItemColsExcelDto;
import vn.com.phat.excel.enumdef.JavaDataType;
import vn.com.phat.excel.exception.ExcelException;
import vn.com.phat.excel.param.ExcelExporterParam;
import vn.com.phat.excel.param.ExcelExporterSheetParam;

import java.io.File;
import java.io.FileInputStream;
import java.io.OutputStream;
import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.Comparator;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Optional;
import java.util.concurrent.ConcurrentHashMap;
import java.util.concurrent.ExecutorService;
import java.util.concurrent.Executors;
import java.util.concurrent.TimeUnit;
import java.util.function.Consumer;
import java.util.function.Supplier;
import java.util.stream.Collectors;

/**
 * Performs the export of data to an Excel file.
 * The OutputStream where the Excel file will be written. This is used to write the Excel file.
 * The parameters for the export. This includes the data to be exported and any additional settings.
 * throws ExcelException If an error occurs during the export.
 * @author phatlt
 */
@Slf4j
public class ExcelExporterDefaultImpl implements ExcelExporter {

    private final Map<String, CellDataTypeHandler> cellDataHandler;
    private DataFormat dataFormat;
    private final Map<String, CellStyle> cellStyleMap;
    private final Map<String, CellStyleHandler> cellStyleHandler;


    // By default, use this constructor is enough.
    public ExcelExporterDefaultImpl(){
        this.cellDataHandler = new HashMap<>();
        this.cellDataHandler.put(JavaDataType.STRING.name(), new CellStringHandler());
        this.cellDataHandler.put(JavaDataType.LONG.name(), new CellLongHandler());
        this.cellDataHandler.put(JavaDataType.DOUBLE.name(), new CellDoubleHandler());
        this.cellDataHandler.put(JavaDataType.INTEGER.name(), new CellIntegerHandler());
        this.cellDataHandler.put(JavaDataType.INT.name(), new CellIntegerHandler());
        this.cellDataHandler.put(JavaDataType.DATE.name(), new CellDateHandler());
        this.cellDataHandler.put(JavaDataType.TIMESTAMP.name(), new CellDateHandler());
        this.cellDataHandler.put(JavaDataType.BIGDECIMAL.name(), new CellBigDecimalHandler());
        this.cellDataHandler.put(JavaDataType.BOOLEAN.name(), new CellBooleanHandler());

        this.cellStyleMap = new ConcurrentHashMap<>();
        this.cellStyleHandler = new HashMap<>();
        this.cellStyleHandler.put(JavaDataType.DATE.name(), new CellStyleDateHandler());
        this.cellStyleHandler.put(JavaDataType.BIGDECIMAL.name(), new CellStyleBigDecimalHandler());
        this.cellStyleHandler.put(JavaDataType.STRING.name(), new CellStyleStringHandler());
        this.cellStyleHandler.put(JavaDataType.BOOLEAN.name(), new CellStyleBooleanHandler());
    }

    // Use this constructor when you want to add more or change behavior of the writing cell data process
    public ExcelExporterDefaultImpl(Map<String, CellDataTypeHandler> dataTypeHandler){
        this();
        Optional.ofNullable(dataTypeHandler).ifPresent(cellDataHandler::putAll);
    }

    public ExcelExporterDefaultImpl(Map<String, CellDataTypeHandler> dataTypeHandler, Map<String, CellStyleHandler> cellStyleHandler){
        this(dataTypeHandler);
        Optional.ofNullable(cellStyleHandler).ifPresent(this.cellStyleHandler::putAll);
    }

    /**
     * Retrieves an existing sheet from the workbook by index if it is within the template sheet count; otherwise, creates a new sheet with the specified name.
     *
     * @param sheetParam parameters containing the sheet index and name
     * @param workbook the streaming workbook instance
     * @param numberOfSheetTemplate the number of sheets present in the template
     * @return the existing or newly created SXSSFSheet
     */
    private SXSSFSheet getSheet(ExcelExporterSheetParam sheetParam, SXSSFWorkbook workbook, int numberOfSheetTemplate){
        return numberOfSheetTemplate > sheetParam.getSheetIndex()
                ? workbook.getSheetAt(sheetParam.getSheetIndex())
                : workbook.createSheet(sheetParam.getSheetName());
    }

    /**
     * Exports data to an Excel file using a template and writes the result to the provided output stream.
     *
     * @param outStream the output stream to write the generated Excel file to
     * @param param the export parameters, including template path and sheet data
     * @throws ExcelException if an error occurs during the export process
     */
    @Override
    public void doExport(OutputStream outStream, ExcelExporterParam param) throws ExcelException {
        Assert.notNull(param, "The parameter for the export must not be null.");
        Assert.notNull(outStream, "The OutputStream for the export must not be null.");
        Assert.notNull(param.getTemplatePath(),"The template path must not be null.");
        File file = new File(param.getTemplatePath());
        Assert.isTrue(file.exists() && file.isFile(), "The template file does not exist.");
        try (
            FileInputStream in = new FileInputStream(file);
            XSSFWorkbook xssfWorkbook = new XSSFWorkbook(in);
            SXSSFWorkbook sxssfWorkbook = new SXSSFWorkbook(xssfWorkbook)) {
            this.dataFormat = Optional.ofNullable(this.dataFormat).orElseGet(sxssfWorkbook::createDataFormat);
            int availableProcessors = Runtime.getRuntime().availableProcessors();
            int exportSheets = Optional.ofNullable(param.getSheets()).orElseGet(ArrayList::new).size();
            int theBestPoolSize = Math.min(exportSheets, availableProcessors + 1);
            ExecutorService executorService = Executors.newFixedThreadPool(theBestPoolSize);
            Consumer<List<ExcelExporterSheetParam>> processCollection = getListConsumer(sxssfWorkbook, executorService);
            Optional.ofNullable(param.getSheets()).ifPresent(processCollection);
            executorService.shutdown();
            boolean isCompleted = executorService.awaitTermination(2, TimeUnit.MINUTES);
            if(!isCompleted)
                log.debug("The writing process has timed out.");
            Optional.ofNullable(param.getConsumer()).ifPresent(consumer -> consumer.accept(sxssfWorkbook));
            sxssfWorkbook.write(outStream);
            sxssfWorkbook.dispose();
            outStream.flush();
        } catch (Exception e) {
            log.error(e.getMessage(), e);
            Thread.currentThread().interrupt();
            throw new ExcelException(e);
        }
    }

    /**
     * Returns a consumer that submits sheet data export tasks to the provided executor service.
     *
     * The consumer sorts the given list of sheet parameters by their sheet index, obtains or creates the corresponding sheets in the workbook, and submits a runnable for each sheet to export its data concurrently.
     *
     * @param sxssfWorkbook the streaming workbook to write data into
     * @param executorService the executor service for concurrent task execution
     * @return a consumer that processes and submits export tasks for a list of sheet parameters
     */
    private Consumer<List<ExcelExporterSheetParam>> getListConsumer(SXSSFWorkbook sxssfWorkbook, ExecutorService executorService) {
        final int numberOfSheetTemplate = sxssfWorkbook.getNumberOfSheets();
        return sheets -> {
            List<ExcelExporterSheetParam> sheetOrder = sheets.stream().sorted(
                    Comparator.comparingInt(ExcelExporterSheetParam::getSheetIndex)).toList();
            for(ExcelExporterSheetParam sheet : sheetOrder){
                SXSSFSheet workSheet = getSheet(sheet, sxssfWorkbook, numberOfSheetTemplate);
                Runnable runnable = () -> setDataToSheet(sxssfWorkbook, workSheet, sheet);
                executorService.submit(runnable);
            }
        };
    }

    protected <T extends ExcelExporterSheetParam> void setDataToSheet(final SXSSFWorkbook workbook, final SXSSFSheet sheet, final T sheetParam){
        try {
            String startCell = sheetParam.getStartCell();
            CellReference cellReference = new CellReference(startCell);
            int startRow = cellReference.getRow();
            List<Map<String, Object>> data = Optional.ofNullable(sheetParam.getData()).orElseGet(ArrayList::new);
            List<ItemColsExcelDto> colIndex = Optional.ofNullable(sheetParam.getColIndex()).orElseGet(ArrayList::new)
                    .stream().sorted(Comparator.comparingInt(ItemColsExcelDto::getColIndex)).collect(Collectors.toList());
            for (int i = 0; i < data.size(); i++) {
                final int rowIndex = i + startRow;
                final Map<String, Object> dat = data.get(i);
                final Supplier<SXSSFRow> createRow = () -> sheet.createRow(rowIndex);
                SXSSFRow row = Optional.ofNullable(sheet.getRow(rowIndex)).orElseGet(createRow);
                setDataRow(workbook, row, dat, colIndex, sheetParam.getMapFields());
            }
            if(sheetParam.getConsumer() != null)
                sheetParam.getConsumer().accept(sheet);
        }
        catch (Exception e){
            log.error(e.getMessage(), e);
        }
    }

    protected void setDataRow(SXSSFWorkbook workbook ,SXSSFRow row, Map<String, Object> data, List<ItemColsExcelDto> colIndex, Map<String, Field> mapFields) {
        for (ItemColsExcelDto col : colIndex) {
            int colIn = col.getColIndex();
            Supplier<SXSSFCell> createCell = () -> row.createCell(colIn);
            SXSSFCell cell = Optional.ofNullable(row.getCell(colIn)).orElseGet(createCell);
            final String colName = col.getColName().toUpperCase();
            Field field = mapFields.get(colName);
            Object value = data.get(colName);
            setCellData(workbook ,cell, value, field, colName);
        }
    }

    protected void setCellData(SXSSFWorkbook workbook ,SXSSFCell cell, Object value, Field field, final String colName) {
        final String fieldClassName = field.getType().getSimpleName().toUpperCase();
        CellStyle cellStyle = Optional.ofNullable(getCellStyle(colName, workbook)).orElseGet(() -> getCellStyle(fieldClassName, workbook));
        if(cellStyle!= null)
            cell.setCellStyle(cellStyle);
        CellDataTypeHandler dataTypeHandler = Optional.ofNullable(this.cellDataHandler.get(colName)).orElseGet(()-> this.cellDataHandler.get(fieldClassName));
        Optional.ofNullable(dataTypeHandler).ifPresent(handler -> handler.setCellData(cell, value));
    }

    protected CellStyle getCellStyle(String name, SXSSFWorkbook workbook){
        Supplier<CellStyle> createCellStyle = () -> {
            CellStyleHandler handler = this.cellStyleHandler.get(name);
            if(handler == null) return null;
            return handler.handleCellStyle(workbook, this.dataFormat);
        };
        CellStyle cellStyle = Optional.ofNullable(cellStyleMap.get(name)).orElseGet(createCellStyle);
        if(!cellStyleMap.containsKey(name) && cellStyle != null)
            cellStyleMap.put(name, cellStyle);

        return cellStyle;
    }
}
