package com.excel;

import cn.hutool.core.bean.BeanUtil;
import cn.hutool.core.collection.CollUtil;
import cn.hutool.core.comparator.IndexedComparator;
import cn.hutool.core.io.FileUtil;
import cn.hutool.core.io.IORuntimeException;
import cn.hutool.core.io.IoUtil;
import cn.hutool.core.lang.Assert;
import cn.hutool.core.map.MapUtil;
import cn.hutool.core.util.CharsetUtil;
import cn.hutool.core.util.IdUtil;
import cn.hutool.core.util.StrUtil;
import cn.hutool.core.util.URLUtil;
import cn.hutool.poi.excel.*;
import cn.hutool.poi.excel.cell.CellLocation;
import cn.hutool.poi.excel.cell.CellUtil;
import cn.hutool.poi.excel.style.Align;
import com.google.common.collect.Lists;
import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddressList;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFDataValidation;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.ServletOutputStream;
import javax.servlet.http.HttpServletResponse;
import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;
import java.lang.reflect.Field;
import java.net.URLEncoder;
import java.nio.charset.Charset;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.*;
import java.util.concurrent.ConcurrentHashMap;
import java.util.concurrent.atomic.AtomicInteger;

/**
 * Hutool excel导出工具封装
 *
 * @author: MingWei Yang
 * @since: 2020/10/29 17:57
 * @description: excel导出工具
 */
public class ExcelUtils extends ExcelUtil {

    /**
     * excel导出名称
     */
    @Target({ElementType.TYPE})
    @Retention(RetentionPolicy.RUNTIME)
    public @interface ExcelName {
        String value() default "";
    }

    /**
     * excel导出Sheet名称
     */
    @Target({ElementType.TYPE})
    @Retention(RetentionPolicy.RUNTIME)
    public @interface SheetName {
        String value() default "";
    }

    /**
     * excel导出列宽
     */
    @Target({ElementType.FIELD})
    @Retention(RetentionPolicy.RUNTIME)
    public @interface ExcelColumn {
        String value() default "";

        int width() default 15;

        int col() default 0;

        boolean required() default false;
    }

    /**
     * excel导出列宽
     */
    @Target({ElementType.FIELD, ElementType.TYPE})
    @Retention(RetentionPolicy.RUNTIME)
    public @interface MergeTitle {
        String value() default "";

        int col() default 1;
    }

    /**
     * excel封装属性对象
     */
    @Data
    @AllArgsConstructor
    @NoArgsConstructor
    public static class Excel<T> {
        /**
         * excel对象Class
         */
        private Class<T> cls;
        /**
         * 数据
         */
        private Collection<T> coll;
        /**
         * 合并单元格
         */
        private List<ExcelUtilsMerge> mergeList;

        public Excel(Class<T> cls, Collection<T> coll) {
            this.cls = cls;
            this.coll = coll;
        }

        public static <T> Excel<T> build(Class<T> cls, Collection<T> coll) {
            Excel<T> excel = new Excel();
            excel.setCls(cls);
            excel.setColl(coll);
            return excel;
        }
    }

    public static class ExcelWriter extends ExcelBase<ExcelWriter> {
        protected File destFile;
        private AtomicInteger currentRow;
        private Map<String, String> headerAlias;
        private boolean onlyAlias;
        private Comparator<String> aliasComparator;
        private StyleSet styleSet;
        private Map<String, Integer> headLocationCache;

        public ExcelWriter() {
            this(false);
        }

        public ExcelWriter(boolean isXlsx) {
            this((Workbook) WorkbookUtil.createBook(isXlsx), (String) null);
        }

        public ExcelWriter(String destFilePath) {
            this((String) destFilePath, (String) null);
        }

        public ExcelWriter(boolean isXlsx, String sheetName) {
            this(WorkbookUtil.createBook(isXlsx), sheetName);
        }

        public ExcelWriter(String destFilePath, String sheetName) {
            this(cn.hutool.core.io.FileUtil.file(destFilePath), sheetName);
        }

        public ExcelWriter(File destFile) {
            this((File) destFile, (String) null);
        }

        public ExcelWriter(File destFile, String sheetName) {
            this(WorkbookUtil.createBookForWriter(destFile), sheetName);
            this.destFile = destFile;
        }

        public ExcelWriter(Workbook workbook, String sheetName) {
            this(WorkbookUtil.getOrCreateSheet(workbook, sheetName));
        }

        public ExcelWriter(Sheet sheet) {
            super(sheet);
            this.currentRow = new AtomicInteger(0);
            this.styleSet = new StyleSet(this.workbook);
        }

        @Override
        public ExcelWriter setSheet(int sheetIndex) {
            this.reset();
            return (ExcelWriter) super.setSheet(sheetIndex);
        }

        @Override
        public ExcelWriter setSheet(String sheetName) {
            this.reset();
            return (ExcelWriter) super.setSheet(sheetName);
        }

        public ExcelWriter reset() {
            this.resetRow();
            this.headLocationCache = null;
            return this;
        }

        public ExcelWriter renameSheet(String sheetName) {
            return this.renameSheet(this.workbook.getSheetIndex(this.sheet), sheetName);
        }

        public ExcelWriter renameSheet(int sheet, String sheetName) {
            this.workbook.setSheetName(sheet, sheetName);
            return this;
        }

        public ExcelWriter autoSizeColumnAll() {
            int columnCount = this.getColumnCount();

            for (int i = 0; i < columnCount; ++i) {
                this.autoSizeColumn(i);
            }

            return this;
        }

        public ExcelWriter autoSizeColumn(int columnIndex) {
            this.sheet.autoSizeColumn(columnIndex);
            return this;
        }

        public ExcelWriter autoSizeColumn(int columnIndex, boolean useMergedCells) {
            this.sheet.autoSizeColumn(columnIndex, useMergedCells);
            return this;
        }

        public ExcelWriter disableDefaultStyle() {
            return this.setStyleSet((StyleSet) null);
        }

        public ExcelWriter setStyleSet(StyleSet styleSet) {
            this.styleSet = styleSet;
            return this;
        }

        public StyleSet getStyleSet() {
            return this.styleSet;
        }

        public CellStyle getHeadCellStyle() {
            return this.styleSet.getHeadCellStyle();
        }

        public CellStyle getCellStyle() {
            return null == this.styleSet ? null : this.styleSet.getCellStyle();
        }

        public int getCurrentRow() {
            return this.currentRow.get();
        }

        public String getDisposition(String fileName, Charset charset) {
            if (null == charset) {
                charset = CharsetUtil.CHARSET_UTF_8;
            }

            if (StrUtil.isBlank(fileName)) {
                fileName = IdUtil.fastSimpleUUID();
            }

            fileName = StrUtil.addSuffixIfNot(URLUtil.encodeAll(fileName, charset), this.isXlsx() ? ".xlsx" : ".xls");
            return StrUtil.format("attachment; filename=\"{}\"; filename*={}''{}", new Object[]{fileName, charset.name(), fileName});
        }

        public ExcelWriter setCurrentRow(int rowIndex) {
            this.currentRow.set(rowIndex);
            return this;
        }

        public ExcelWriter setCurrentRowToEnd() {
            return this.setCurrentRow(this.getRowCount());
        }

        public ExcelWriter passCurrentRow() {
            this.currentRow.incrementAndGet();
            return this;
        }

        public ExcelWriter passRows(int rows) {
            this.currentRow.addAndGet(rows);
            return this;
        }

        public ExcelWriter resetRow() {
            this.currentRow.set(0);
            return this;
        }

        public ExcelWriter setDestFile(File destFile) {
            this.destFile = destFile;
            return this;
        }

        public ExcelWriter setHeaderAlias(Map<String, String> headerAlias) {
            this.headerAlias = headerAlias;
            this.aliasComparator = null;
            return this;
        }

        public ExcelWriter clearHeaderAlias() {
            this.headerAlias = null;
            this.aliasComparator = null;
            return this;
        }

        public ExcelWriter setOnlyAlias(boolean isOnlyAlias) {
            this.onlyAlias = isOnlyAlias;
            return this;
        }

        public ExcelWriter addHeaderAlias(String name, String alias) {
            Map<String, String> headerAlias = this.headerAlias;
            if (null == headerAlias) {
                headerAlias = new LinkedHashMap();
            }

            this.headerAlias = (Map) headerAlias;
            ((Map) headerAlias).put(name, alias);
            this.aliasComparator = null;
            return this;
        }

        public ExcelWriter setFreezePane(int rowSplit) {
            return this.setFreezePane(0, rowSplit);
        }

        public ExcelWriter setFreezePane(int colSplit, int rowSplit) {
            this.getSheet().createFreezePane(colSplit, rowSplit);
            return this;
        }

        public ExcelWriter setColumnWidth(int columnIndex, int width) {
            if (columnIndex < 0) {
                this.sheet.setDefaultColumnWidth(width);
            } else {
                this.sheet.setColumnWidth(columnIndex, width * 256);
            }

            return this;
        }

        public ExcelWriter setDefaultRowHeight(int height) {
            return this.setRowHeight(-1, height);
        }

        public ExcelWriter setRowHeight(int rownum, int height) {
            if (rownum < 0) {
                this.sheet.setDefaultRowHeightInPoints((float) height);
            } else {
                Row row = this.sheet.getRow(rownum);
                if (null != row) {
                    row.setHeightInPoints((float) height);
                }
            }

            return this;
        }

        public ExcelWriter setHeaderOrFooter(String text, Align align, boolean isFooter) {
            HeaderFooter headerFooter = isFooter ? this.sheet.getFooter() : this.sheet.getHeader();
            switch (align) {
                case LEFT:
                    ((HeaderFooter) headerFooter).setLeft(text);
                    break;
                case RIGHT:
                    ((HeaderFooter) headerFooter).setRight(text);
                    break;
                case CENTER:
                    ((HeaderFooter) headerFooter).setCenter(text);
            }

            return this;
        }

        public ExcelWriter addSelect(int x, int y, String... selectList) {
            return this.addSelect(new CellRangeAddressList(y, y, x, x), selectList);
        }

        public ExcelWriter addSelect(CellRangeAddressList regions, String... selectList) {
            DataValidationHelper validationHelper = this.sheet.getDataValidationHelper();
            DataValidationConstraint constraint = validationHelper.createExplicitListConstraint(selectList);
            DataValidation dataValidation = validationHelper.createValidation(constraint, regions);
            if (dataValidation instanceof XSSFDataValidation) {
                dataValidation.setSuppressDropDownArrow(true);
                dataValidation.setShowErrorBox(true);
            } else {
                dataValidation.setSuppressDropDownArrow(false);
            }

            return this.addValidationData(dataValidation);
        }

        public ExcelWriter addValidationData(DataValidation dataValidation) {
            this.sheet.addValidationData(dataValidation);
            return this;
        }

        public ExcelWriter merge(int lastColumn) {
            return this.merge(lastColumn, (Object) null);
        }

        public ExcelWriter merge(int lastColumn, Object content) {
            return this.merge(lastColumn, content, true);
        }

        public ExcelWriter merge(int lastColumn, Object content, boolean isSetHeaderStyle) {
            Assert.isFalse(this.isClosed, "ExcelWriter has been closed!", new Object[0]);
            int rowIndex = this.currentRow.get();
            this.merge(rowIndex, rowIndex, 0, lastColumn, content, isSetHeaderStyle);
            if (null != content) {
                this.currentRow.incrementAndGet();
            }

            return this;
        }

        public ExcelWriter merge(int firstRow, int lastRow, int firstColumn, int lastColumn, Object content, boolean isSetHeaderStyle) {
            Assert.isFalse(this.isClosed, "ExcelWriter has been closed!", new Object[0]);
            CellStyle style = null;
            if (null != this.styleSet) {
                style = isSetHeaderStyle && null != this.styleSet.getHeadCellStyle() ? this.styleSet.getHeadCellStyle() : this.styleSet.getCellStyle();
            }

            CellUtil.mergingCells(this.sheet, firstRow, lastRow, firstColumn, lastColumn, style);
            if (null != content) {
                Cell cell = this.getOrCreateCell(firstColumn, firstRow);
                CellUtil.setCellValue(cell, content, this.styleSet, isSetHeaderStyle);
            }

            return this;
        }

        public ExcelWriter write(Iterable<?> data) {
            return this.write(data, 0 == this.getCurrentRow());
        }

        public ExcelWriter write(Iterable<?> data, boolean isWriteKeyAsHead) {
            Assert.isFalse(this.isClosed, "ExcelWriter has been closed!", new Object[0]);
            boolean isFirst = true;
            Iterator var4 = data.iterator();

            while (var4.hasNext()) {
                Object object = var4.next();
                this.writeRow(object, isFirst && isWriteKeyAsHead);
                if (isFirst) {
                    isFirst = false;
                }
            }

            return this;
        }

        public ExcelWriter write(Iterable<?> data, Comparator<String> comparator) {
            Assert.isFalse(this.isClosed, "ExcelWriter has been closed!", new Object[0]);
            boolean isFirstRow = true;
            Iterator var5 = data.iterator();

            while (var5.hasNext()) {
                Object obj = var5.next();
                Object map;
                if (obj instanceof Map) {
                    map = new TreeMap(comparator);
                    ((Map) map).putAll((Map) obj);
                } else {
                    map = BeanUtil.beanToMap(obj, new TreeMap(comparator), false, false);
                }

                this.writeRow((Map) map, isFirstRow);
                if (isFirstRow) {
                    isFirstRow = false;
                }
            }

            return this;
        }

        public ExcelWriter writeHeadRow() {
            Assert.isFalse(this.isClosed, "ExcelWriter has been closed!", new Object[0]);
            this.headLocationCache = new ConcurrentHashMap();
            Row row = this.sheet.createRow(this.currentRow.getAndIncrement());

            int i = 0;
            for (Iterator var5 = this.headerAlias.entrySet().iterator(); var5.hasNext(); ++i) {
                Map.Entry<?, ?> header = (Map.Entry) var5.next();
                Cell cell = row.createCell(i);
                if (null != header.getValue()) {
                    CellUtil.setCellValue(cell, header.getValue(), this.styleSet, true);
                    this.headLocationCache.put(StrUtil.toString(header.getKey()), i);
                } else if (!this.onlyAlias) {
                    CellUtil.setCellValue(cell, header.getKey(), this.styleSet, true);
                    this.headLocationCache.put(StrUtil.toString(header.getKey()), i);
                }
            }

            return this;
        }

        public ExcelWriter writeRow(Object rowBean, boolean isWriteKeyAsHead) {
            if (rowBean instanceof Iterable) {
                return this.writeRow((Iterable) rowBean);
            } else {
                Object rowMap;
                if (rowBean instanceof Map) {
                    if (MapUtil.isNotEmpty(this.headerAlias)) {
                        rowMap = MapUtil.newTreeMap((Map) rowBean, this.getCachedAliasComparator());
                    } else {
                        rowMap = (Map) rowBean;
                    }
                } else {
                    if (!BeanUtil.isBean(rowBean.getClass())) {
                        return this.writeRow((Object) CollUtil.newArrayList(new Object[]{rowBean}), isWriteKeyAsHead);
                    }

                    if (MapUtil.isEmpty(this.headerAlias)) {
                        rowMap = BeanUtil.beanToMap(rowBean, new LinkedHashMap(), false, false);
                    } else {
                        rowMap = BeanUtil.beanToMap(rowBean, new TreeMap(this.getCachedAliasComparator()), false, false);
                    }
                }

                return this.writeRow((Map) rowMap, isWriteKeyAsHead);
            }
        }

        public ExcelWriter writeRow(Map<?, ?> rowMap, boolean isWriteKeyAsHead) {
            Assert.isFalse(this.isClosed, "ExcelWriter has been closed!", new Object[0]);
            if (MapUtil.isEmpty(rowMap)) {
                return this.passCurrentRow();
            } else {
                if (isWriteKeyAsHead) {
                    this.writeHeadRow();
                }

                if (MapUtil.isNotEmpty(this.headLocationCache)) {
                    Row row = RowUtil.getOrCreateRow(this.sheet, this.currentRow.getAndIncrement());
                    Iterator var6 = rowMap.entrySet().iterator();

                    while (var6.hasNext()) {
                        Map.Entry<?, ?> entry = (Map.Entry) var6.next();
                        Integer location = this.headLocationCache.get(StrUtil.toString(entry.getKey()));
                        if (null != location) {
                            CellUtil.setCellValue(CellUtil.getOrCreateCell(row, location), entry.getValue(), this.styleSet, false);
                        }
                    }
                } else {
                    this.writeRow(rowMap.values());
                }

                return this;
            }
        }

        public ExcelWriter writeRow(Iterable<?> rowData) {
            Assert.isFalse(this.isClosed, "ExcelWriter has been closed!", new Object[0]);
            RowUtil.writeRow(this.sheet.createRow(this.currentRow.getAndIncrement()), rowData, this.styleSet, false);
            return this;
        }

        public ExcelWriter writeCellValue(String locationRef, Object value) {
            CellLocation cellLocation = ExcelUtil.toLocation(locationRef);
            return this.writeCellValue(cellLocation.getX(), cellLocation.getY(), value);
        }

        public ExcelWriter writeCellValue(int x, int y, Object value) {
            Cell cell = this.getOrCreateCell(x, y);
            CellUtil.setCellValue(cell, value, this.styleSet, false);
            return this;
        }

        /**
         * @deprecated
         */
        @Deprecated
        public CellStyle createStyleForCell(int x, int y) {
            return this.createCellStyle(x, y);
        }

        public ExcelWriter setStyle(CellStyle style, String locationRef) {
            CellLocation cellLocation = ExcelUtil.toLocation(locationRef);
            return this.setStyle(style, cellLocation.getX(), cellLocation.getY());
        }

        public ExcelWriter setStyle(CellStyle style, int x, int y) {
            Cell cell = this.getOrCreateCell(x, y);
            cell.setCellStyle(style);
            return this;
        }

        public ExcelWriter setRowStyle(int y, CellStyle style) {
            this.getOrCreateRow(y).setRowStyle(style);
            return this;
        }

        public Font createFont() {
            return this.getWorkbook().createFont();
        }

        public ExcelWriter flush() throws IORuntimeException {
            return this.flush(this.destFile);
        }

        public ExcelWriter flush(File destFile) throws IORuntimeException {
            Assert.notNull(destFile, "[destFile] is null, and you must call setDestFile(File) first or call flush(OutputStream).", new Object[0]);
            return this.flush(FileUtil.getOutputStream(destFile), true);
        }

        public ExcelWriter flush(OutputStream out) throws IORuntimeException {
            return this.flush(out, false);
        }

        public ExcelWriter flush(OutputStream out, boolean isCloseOut) throws IORuntimeException {
            Assert.isFalse(this.isClosed, "ExcelWriter has been closed!", new Object[0]);

            try {
                this.workbook.write(out);
                out.flush();
            } catch (IOException var7) {
                throw new IORuntimeException(var7);
            } finally {
                if (isCloseOut) {
                    IoUtil.close(out);
                }

            }

            return this;
        }

        @Override
        public void close() {
            if (null != this.destFile) {
                this.flush();
            }

            this.closeWithoutFlush();
        }

        protected void closeWithoutFlush() {
            super.close();
            this.currentRow = null;
            this.styleSet = null;
        }

        private Comparator<String> getCachedAliasComparator() {
            if (MapUtil.isEmpty(this.headerAlias)) {
                return null;
            } else {
                Comparator<String> aliasComparator = this.aliasComparator;
                if (null == aliasComparator) {
                    Set<String> keySet = this.headerAlias.keySet();
                    aliasComparator = new IndexedComparator(keySet.toArray(new String[0]));
                    this.aliasComparator = (Comparator) aliasComparator;
                }

                return (Comparator) aliasComparator;
            }
        }
    }


    /**
     * 封装Hutool导出
     *
     * @param response
     * @param collection
     * @param cls
     * @param <T>
     */
    public static <T> void exportExcel(HttpServletResponse response, Collection<T> collection, Class<T> cls, List<ExcelUtilsMerge> mergeList) {
        exportExcel(response, new Excel(cls, collection, mergeList));
    }

    /**
     * 封装Hutool导出
     *
     * @param response
     * @param collection
     * @param cls
     * @param <T>
     */
    public static <T> void exportExcel(HttpServletResponse response, Collection<T> collection, Class<T> cls) {
        exportExcel(response, new Excel(cls, collection));
    }

    /**
     * 封装Hutool导出
     *
     * @param response
     * @param excels
     */
    public static <T> void exportExcel(HttpServletResponse response, Excel... excels) {
        if (ObjectUtil.isEmpty(excels)) {
            throw new RuntimeException("导出Excel错误，请检查传入Excel信息是否正确");
        }
        String excelName = "";
        for (Excel excel : excels) {
            Class<T> cls = excel.getCls();
            if (ObjectUtil.isNotEmpty(cls)) {
                ExcelName name = cls.getAnnotation(ExcelName.class);
                if (ObjectUtil.isNotEmpty(name)) {
                    excelName = name.value();
                }
                break;
            }
        }
        exportExcel(response, excelName, excels);
    }

    /**
     * 封装Hutool导出
     *
     * @param response
     * @param excels
     */
    public static void exportExcel(HttpServletResponse response, String excelName, Excel... excels) {
        ServletOutputStream outputStream = null;
        ExcelWriter writer = null;
        SXSSFWorkbook workbook = null;
        try {
            workbook = new SXSSFWorkbook();
            // 临时文件将被gzip压缩
            workbook.setCompressTempFiles(true);
            writer = convertWriter(workbook, excelName, excels);
            response.setHeader("Content-disposition", "attachment; filename=" + URLEncoder.encode(excelName + "_" + LocalDateTime.now().format(DateTimeFormatter.ofPattern("yyyy-MM-dd_HH-mm-ss")), StringPool.UTF_8) + ".xlsx");
            outputStream = response.getOutputStream();
            writer.flush(outputStream, true);
        } catch (Exception ex) {
            ex.printStackTrace();
        } finally {
            if (writer != null) {
                writer.close();
            }
            if (outputStream != null) {
                IoUtil.close(outputStream);
            }
            if (workbook != null) {
                //删除磁盘上临时文件
                workbook.dispose();
            }
        }
    }

    /**
     * 转换Excel
     *
     * @param excelName
     * @param excels
     * @return
     */
    private static ExcelWriter convertWriter(SXSSFWorkbook workbook, String excelName, Excel[] excels) {
        ExcelWriter writer = null;
        for (Excel excel : excels) {
            writer = initExcel(workbook, writer, excelName, excel);
        }
        return writer;
    }

    /**
     * 加载Excel内容
     *
     * @param writer
     * @param excelName
     * @param excel
     */
    public static ExcelWriter initExcel(SXSSFWorkbook workbook, ExcelWriter writer, String excelName, Excel excel) {
        //创建Sheet
        Sheet sheet = createSheet(workbook, excel, excelName);
        if (ObjectUtil.isEmpty(writer)) {
            writer = new ExcelWriter(sheet);
        } else {
            //设置Sheet名称
            writer.setSheet(sheet);
        }
        Field[] fields = excel.getCls().getDeclaredFields();
        //加载Sheet自定义标题别名
        initSheetTitle(writer, fields);
        //加载数据（校验空数据）
        initNullColl(excel);
        //重置在第一行写入数据
        writer.reset();
        //加载合并单元格标题行
        initMergeTitle(writer, fields, excel.getCls());
        //写入数据
        writer.write(excel.getColl(), true);
        //加载合并单元格
        List<ExcelUtilsMerge> mergeList = excel.getMergeList();
        if (ObjectUtil.isNotEmpty(mergeList)) {
            for (ExcelUtilsMerge merge : mergeList) {
                writer.merge(merge.getFirstRow(), merge.getLastRow(), merge.getFirstColumn(), merge.getLastColumn(), merge.getContent(), merge.isSetHeaderStyle());
            }
        }
        return writer;
    }

    /**
     * 加载合并单元格后的标题行
     *
     * @param writer
     * @param fields
     */
    private static <T> void initMergeTitle(ExcelWriter writer, Field[] fields, Class<T> cls) {
        MergeTitle validateMergeTitle = cls.getAnnotation(MergeTitle.class);
        if (null == validateMergeTitle) {
            return;
        }
        for (int i = 0; i < fields.length; i++) {
            Field field = fields[i];
            MergeTitle mergeTitle = field.getAnnotation(MergeTitle.class);
            if (null != mergeTitle) {
                if (mergeTitle.col() > 1) {
                    writer.merge(0, 0, i, i + mergeTitle.col() - 1, mergeTitle.value(), true);
                } else {
                    Cell cell = writer.getOrCreateCell(i + mergeTitle.col() - 1, 0);
                    CellUtil.setCellValue(cell, mergeTitle.value(), writer.getStyleSet(), true);
                }
            }
        }
        writer.passCurrentRow();
    }

    /**
     * 加载Sheet自定义标题别名
     *
     * @param writer
     * @param fields
     */
    private static void initSheetTitle(ExcelWriter writer, Field[] fields) {
        for (int i = 0; i < fields.length; i++) {
            Field field = fields[i];
            ExcelColumn column = field.getAnnotation(ExcelColumn.class);
            if (ObjectUtil.isNotEmpty(column)) {
                //标题名称
                if (ObjectUtil.isNotEmpty(column.value())) {
                    writer.addHeaderAlias(field.getName(), column.value());
                } else {
                    writer.addHeaderAlias(field.getName(), null);
                }
                //列宽
                if (ObjectUtil.isNotEmpty(column.width())) {
                    writer.setColumnWidth(i, column.width());
                }
            }
        }
        //默认的，未添加alias的属性也会写出，如果想只写出加了别名的字段，可以调用此方法排除之
        writer.setOnlyAlias(true);
    }

    /**
     * 创建Sheet
     *
     * @param workbook
     * @param excel
     * @param excelName
     * @return
     */
    private static Sheet createSheet(SXSSFWorkbook workbook, Excel excel, String excelName) {
        SXSSFSheet sheet = workbook.createSheet(getSheetName(excel.getCls(), excelName));
        if (ObjectUtil.isNotEmpty(excel.getMergeList())) {
            sheet.setRandomAccessWindowSize(-1);
        } else {
            sheet.setRandomAccessWindowSize(500);
        }
        return sheet;
    }

    /**
     * 加载一条空数据，防止没有数据下没有表头
     *
     * @param excel
     */
    private static void initNullColl(Excel excel) {
        Collection coll = excel.getColl();
        if (ObjectUtil.isEmpty(coll)) {
            try {
                excel.setColl(Lists.newArrayList(excel.getCls().newInstance()));
            } catch (Exception ex) {
                ex.printStackTrace();
            }
        }
    }

    /**
     * 获取SheetName
     *
     * @param cls
     * @param defaultName
     * @param <T>
     * @return
     */
    private static <T> String getSheetName(Class<T> cls, String defaultName) {
        SheetName name = cls.getAnnotation(SheetName.class);
        if (ObjectUtil.isNotEmpty(name)) {
            return name.value();
        }
        return defaultName;
    }

    /**
     * 读取Excel
     *
     * @param fileData
     * @param cls
     * @return java.util.List<T>
     * @author Mingwei Yang
     * @date 2021-05-11 21:13
     */
    public static <T> List<T> read(MultipartFile fileData, Class<T> cls) {
        return read(fileData, cls, true, 0);
    }

    /**
     * 读取Excel
     *
     * @param fileData
     * @param cls
     * @param isRemoveHead
     * @param headRowIndex
     * @return java.util.List<T>
     * @author Mingwei Yang
     * @date 2021-05-11 21:13
     */
    public static <T> List<T> read(MultipartFile fileData, Class<T> cls, boolean isRemoveHead, int headRowIndex) {
        List<T> listData = new ArrayList<>();
        try {
            InputStream inputStream = fileData.getInputStream();
            ExcelReader excelReader = ExcelUtils.getReader(inputStream);
            List<List<Object>> rowList = excelReader.read();
            if (isRemoveHead && ObjectUtil.isNotEmpty(rowList)) {
                rowList.remove(headRowIndex);
            }
            ForEachUtils.forEach(
                    0, rowList,
                    (rowIndex, row) -> {
                        T obj = newInstance(cls);
                        ForEachUtils.forEach(0, row, (colIndex, col) -> setProperty(rowIndex, colIndex, col, obj));
                        listData.add(obj);
                    }
            );
            if (ObjectUtil.isEmpty(listData)) {
                throw new RuntimeException("暂未读取到Excel数据，请检查Excel是否为空");
            }
        } catch (Exception ex) {
            throw new RuntimeException(ex.getMessage());
        }
        return listData;
    }

    /**
     * 创建对象
     *
     * @param cls
     * @param <T>
     * @return
     */
    private static <T> T newInstance(Class<T> cls) {
        try {
            return cls.newInstance();
        } catch (Exception ex) {
            ex.printStackTrace();
        }
        return null;
    }

    /**
     * 设置属性
     *
     * @param colIndex
     * @param value
     * @param obj
     */
    private static void setProperty(int rowIndex, int colIndex, Object value, Object obj) {
        Field[] fields = obj.getClass().getDeclaredFields();
        for (Field field : fields) {
            ExcelColumn column = field.getAnnotation(ExcelColumn.class);
            if (ObjectUtil.isNotEmpty(column) && ObjectUtil.isNotEmpty(column.col()) && colIndex == column.col()) {
                if (ObjectUtil.isNotEmpty(value)) {
                    try {
                        BeanUtil.setFieldValue(obj, field.getName(), value);
                    } catch (Exception e) {
                        throw new RuntimeException("第" + (rowIndex + 2) + "行：" + column.value() + "字段读取错误，请检查文件内容是否正确 " + value);
                    }
                } else if (column.required()) {
                    throw new RuntimeException(column.value() + "必填");
                }
            }
        }
    }

    /**
     * ExcelWriter关闭
     *
     * @param writer
     * @param outputStream
     * @param workbook
     * @return void
     * @author yang
     * @date 2021/12/30 8:59
     */
    public static void close(ExcelWriter writer, OutputStream outputStream, SXSSFWorkbook workbook) {
        if (writer != null) {
            writer.close();
        }
        if (outputStream != null) {
            IoUtil.close(outputStream);
        }
        if (workbook != null) {
            //删除磁盘上临时文件
            workbook.dispose();
        }
    }

}