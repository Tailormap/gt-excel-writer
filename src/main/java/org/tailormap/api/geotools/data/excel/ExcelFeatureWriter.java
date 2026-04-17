/*
 * Copyright (C) 2026 B3Partners B.V.
 *
 * SPDX-License-Identifier: LGPL-2.1-or-later
 */
package org.tailormap.api.geotools.data.excel;

import java.io.IOException;
import java.util.Calendar;
import java.util.Date;
import java.util.concurrent.atomic.AtomicInteger;
import java.util.logging.Logger;
import org.apache.poi.ss.SpreadsheetVersion;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.geotools.api.data.FeatureWriter;
import org.geotools.api.data.Query;
import org.geotools.api.data.Transaction;
import org.geotools.api.feature.simple.SimpleFeature;
import org.geotools.api.feature.simple.SimpleFeatureType;
import org.geotools.api.feature.type.AttributeDescriptor;
import org.geotools.data.store.ContentEntry;
import org.geotools.data.store.ContentState;
import org.geotools.feature.simple.SimpleFeatureBuilder;
import org.geotools.util.logging.Logging;

public class ExcelFeatureWriter implements FeatureWriter<SimpleFeatureType, SimpleFeature> {
    private static final Logger logger = Logging.getLogger(ExcelFeatureWriter.class);

    private final ContentState state;
    private final SimpleFeatureBuilder featureBuilder;

    private final ExcelDataStore dataStore;
    private final SXSSFSheet sheet;
    private SimpleFeature currentFeature;
    private int nextRow = 0;

    public ExcelFeatureWriter(ContentEntry entry, Query query) throws IOException {
        this.state = entry.getState(Transaction.AUTO_COMMIT);
        this.dataStore = (ExcelDataStore) entry.getDataStore();
        this.featureBuilder = new SimpleFeatureBuilder(this.dataStore.getSchema());

        @SuppressWarnings("PMD.CloseResource") // the workbook is managed in the store
        final SXSSFWorkbook workbook = this.dataStore.getWorkbook();
        this.sheet = workbook.createSheet(entry.getTypeName());

        final Row header = this.sheet.createRow(nextRow++);
        header.setRowStyle(getHeaderStyle(workbook));
        final AtomicInteger colNum = new AtomicInteger(0);
        this.dataStore.getSchema().getAttributeDescriptors().stream()
                .map(AttributeDescriptor::getLocalName)
                .forEach(colName -> {
                    header.createCell(colNum.getAndIncrement()).setCellValue(colName);
                });
    }

    @Override
    public SimpleFeatureType getFeatureType() {
        try {
            return this.dataStore.getSchema();
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }

    @Override
    public boolean hasNext() {
        // append-only writer
        return false;
    }

    @Override
    public SimpleFeature next() throws IOException {
        this.currentFeature = featureBuilder.buildFeature(null);
        return this.currentFeature;
    }

    @Override
    public void remove() throws IOException {
        this.currentFeature = null;
    }

    @Override
    public void write() throws IOException {
        if (this.currentFeature == null) {
            return;
        }
        this.write(currentFeature);
        this.currentFeature = null;
    }

    @Override
    public void close() {
        this.state.close();
    }

    /**
     * Write the next row in the Excel sheet.
     *
     * @param feature to write
     */
    private void write(SimpleFeature feature) {
        final Row row = sheet.createRow(nextRow++);
        // for each attribute create a cell and set the value
        for (int i = 0; i < feature.getAttributeCount(); i++) {
            Object value = feature.getAttribute(i);
            if (value != null) {
                if (value instanceof Number n) {
                    row.createCell(i).setCellValue(n.doubleValue());
                } else if (value instanceof Boolean b) {
                    row.createCell(i).setCellValue(b);
                } else if (value instanceof Date d) {
                    Cell cell = row.createCell(i);
                    cell.setCellValue(d);
                    cell.setCellStyle(getDateStyle(this.dataStore.getWorkbook()));
                } else if (value instanceof Calendar c) {
                    Cell cell = row.createCell(i);
                    cell.setCellValue(c);
                    cell.setCellStyle(getDateStyle(this.dataStore.getWorkbook()));
                } else {
                    if (value.toString().length() > SpreadsheetVersion.EXCEL2007.getMaxTextLength()) {
                        // Excel has a maximum text length; if the value exceeds this, we truncate it and log a
                        // warning
                        // TODO the FID here may not correspond to a FID from the source data
                        logger.warning("Value for attribute '"
                                + feature.getType().getDescriptor(i).getLocalName() + "' on feature '"
                                + feature.getID() + "' exceeds Excel's maximum text length and will be omitted.");

                        // to truncate: value = value.toString().substring(0,
                        //    SpreadsheetVersion.EXCEL2007.getMaxTextLength());
                        // TODO set a warning message instead?
                        row.createCell(i).setBlank();
                    } else {
                        row.createCell(i).setCellValue(value.toString());
                    }
                }
            } else {
                row.createCell(i).setBlank();
            }
        }
        // this could get expensive for large sets!
        this.sheet.trackAllColumnsForAutoSizing();
        row.forEach(item -> {
            this.sheet.autoSizeColumn(item.getColumnIndex());
        });
    }

    private CellStyle getHeaderStyle(Workbook workbook) {
        final CellStyle headerStyle = workbook.createCellStyle();
        Font headerFont = workbook.createFont();
        headerFont.setBold(true);
        headerStyle.setFont(headerFont);
        return headerStyle;
    }

    private CellStyle getDateStyle(Workbook workbook) {
        final CellStyle dateStyle = workbook.createCellStyle();
        CreationHelper createHelper = workbook.getCreationHelper();
        dateStyle.setDataFormat(createHelper.createDataFormat().getFormat("yyyy-mm-dd hh:mm:ss"));
        return dateStyle;
    }
}
