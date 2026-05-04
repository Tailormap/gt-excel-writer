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
import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.Comment;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
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
import org.locationtech.jts.geom.Geometry;

public class ExcelFeatureWriter implements FeatureWriter<SimpleFeatureType, SimpleFeature> {
    private static final Logger logger = Logging.getLogger(ExcelFeatureWriter.class);

    private final ContentState state;
    private final SimpleFeatureBuilder featureBuilder;

    private final ExcelDataStore dataStore;
    private final SXSSFSheet sheet;
    private final CellStyle dateStyle;
    private final CellStyle dateTimeStyle;
    private final CellStyle errorStyle;
    private SimpleFeature currentFeature;
    private int nextRow = 0;
    private final CreationHelper creationHelper;

    public ExcelFeatureWriter(ContentEntry entry, Query query) throws IOException {
        this.state = entry.getState(Transaction.AUTO_COMMIT);
        this.dataStore = (ExcelDataStore) entry.getDataStore();
        this.featureBuilder = new SimpleFeatureBuilder(this.dataStore.getSchema());

        @SuppressWarnings("PMD.CloseResource") // the workbook is managed in the store
        final SXSSFWorkbook workbook = this.dataStore.getWorkbook();
        this.sheet = workbook.createSheet(entry.getTypeName());
        this.dateStyle = createDateStyle(workbook);
        this.dateTimeStyle = createDateTimeStyle(workbook);
        this.errorStyle = createErrorStyle(workbook);
        this.creationHelper = workbook.getCreationHelper();

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
                switch (value) {
                    case Number n -> row.createCell(i).setCellValue(n.doubleValue());
                    case Boolean b -> row.createCell(i).setCellValue(b);
                    case Date d -> {
                        Cell cell = row.createCell(i);
                        cell.setCellValue(d);
                        if (hasTime(d)) {
                            cell.setCellStyle(this.dateTimeStyle);
                        } else {
                            cell.setCellStyle(this.dateStyle);
                        }
                    }
                    case Calendar c -> {
                        Cell cell = row.createCell(i);
                        cell.setCellValue(c);
                        if (hasTime(c)) {
                            cell.setCellStyle(this.dateTimeStyle);
                        } else {
                            cell.setCellStyle(this.dateStyle);
                        }
                    }
                    default -> {
                        Cell cell = row.createCell(i);
                        // handle maximum text length for Excel cells
                        if (value.toString().length() > SpreadsheetVersion.EXCEL2007.getMaxTextLength()) {
                            // TODO the FID here may not correspond to a FID from the source data, but logging the
                            //    offending out-of-bound value is not a good idea
                            logger.info("Value for attribute '"
                                    + feature.getType().getDescriptor(i).getLocalName() + "' on feature '"
                                    + feature.getID() + "' exceeds Excel's maximum text length and will be "
                                    + (value instanceof Geometry ? "omitted." : "truncated."));
                            if (value instanceof Geometry) {
                                // no point is having truncated WKT
                                cell.setBlank();
                            } else {
                                // truncate with ellipsis
                                cell.setCellValue(value.toString()
                                                .substring(0, SpreadsheetVersion.EXCEL2007.getMaxTextLength() - 4)
                                        + " ...");
                            }
                            ClientAnchor anchor = creationHelper.createClientAnchor();
                            anchor.setCol1(cell.getColumnIndex());
                            anchor.setCol2(cell.getColumnIndex() + 3); // comment box width
                            anchor.setRow1(cell.getRowIndex());
                            anchor.setRow2(cell.getRowIndex() + 3); // comment box height
                            Comment comment = sheet.createDrawingPatriarch().createCellComment(anchor);
                            comment.setString(creationHelper.createRichTextString(
                                    "Source value greater than maximum allowed for Excel"));
                            comment.setAddress(cell.getRowIndex(), cell.getColumnIndex());
                            cell.setCellComment(comment);
                            cell.setCellStyle(this.errorStyle);
                        } else {
                            cell.setCellValue(value.toString());
                        }
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

    private boolean hasTime(Date date) {
        Calendar cal = Calendar.getInstance();
        cal.setTime(date);
        return this.hasTime(cal);
    }

    private boolean hasTime(Calendar cal) {
        return cal.get(Calendar.HOUR_OF_DAY) != 0
                || cal.get(Calendar.MINUTE) != 0
                || cal.get(Calendar.SECOND) != 0
                || cal.get(Calendar.MILLISECOND) != 0;
    }

    private CellStyle getHeaderStyle(Workbook workbook) {
        final CellStyle headerStyle = workbook.createCellStyle();
        Font headerFont = workbook.createFont();
        headerFont.setBold(true);
        headerFont.setFontHeightInPoints((short) (workbook.getFontAt((short) 0).getFontHeightInPoints() + 2));
        headerStyle.setFont(headerFont);
        return headerStyle;
    }

    private CellStyle createDateStyle(Workbook workbook) {
        final CellStyle dateStyle = workbook.createCellStyle();
        dateStyle.setDataFormat(workbook.getCreationHelper().createDataFormat().getFormat("yyyy-mm-dd"));
        return dateStyle;
    }

    private CellStyle createDateTimeStyle(Workbook workbook) {
        final CellStyle dateStyle = workbook.createCellStyle();
        dateStyle.setDataFormat(workbook.getCreationHelper().createDataFormat().getFormat("yyyy-mm-dd hh:mm:ss"));
        return dateStyle;
    }

    private CellStyle createErrorStyle(Workbook workbook) {
        final CellStyle errStyle = workbook.createCellStyle();
        Font font = workbook.createFont();
        font.setColor(IndexedColors.RED.getIndex());
        errStyle.setFont(font);
        return errStyle;
    }
}
