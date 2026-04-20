/*
 * Copyright (C) 2026 B3Partners B.V.
 *
 * SPDX-License-Identifier: LGPL-2.1-or-later
 */
package org.tailormap.api.geotools.data.excel;

import static org.hamcrest.MatcherAssert.assertThat;
import static org.hamcrest.Matchers.instanceOf;
import static org.junit.jupiter.api.Assertions.assertAll;
import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertNotNull;
import static org.junit.jupiter.api.Assertions.assertTrue;
import static org.junit.jupiter.api.Assertions.fail;
import static org.junit.jupiter.api.Assumptions.assumeTrue;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Path;
import java.time.LocalDate;
import java.util.Date;
import java.util.List;
import java.util.Map;
import java.util.Random;
import java.util.UUID;
import java.util.concurrent.atomic.AtomicInteger;
import java.util.logging.Level;
import java.util.logging.Logger;
import java.util.stream.IntStream;
import org.apache.poi.ss.SpreadsheetVersion;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.geotools.api.data.FeatureEvent;
import org.geotools.api.data.Query;
import org.geotools.api.data.SimpleFeatureSource;
import org.geotools.api.data.SimpleFeatureStore;
import org.geotools.api.data.Transaction;
import org.geotools.api.feature.simple.SimpleFeatureType;
import org.geotools.api.filter.identity.FeatureId;
import org.geotools.data.DataUtilities;
import org.geotools.data.DefaultTransaction;
import org.geotools.data.memory.MemoryDataStore;
import org.geotools.feature.SchemaException;
import org.geotools.feature.simple.SimpleFeatureBuilder;
import org.geotools.util.logging.Logging;
import org.junit.jupiter.api.AfterEach;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;
import org.locationtech.jts.geom.Coordinate;
import org.locationtech.jts.geom.GeometryFactory;

class ExcelDataStoreFactoryTest {
    private static final Logger logger = Logging.getLogger(ExcelDataStoreFactoryTest.class);
    private static final String FT_NAME = "random";
    private static final int randomFeatureCount = 104;
    private final Date dateWithoutTime = java.sql.Date.valueOf(LocalDate.of(2026, 11, 13));
    private SimpleFeatureSource randomFeatureSource;
    private File excelFile;

    @AfterEach
    void tearDown() {
        randomFeatureSource.getDataStore().dispose();
        if (!logger.isLoggable(Level.FINE)) {
            // the output file is left in the tmp directory for inspection
            excelFile.deleteOnExit();
        }
        randomFeatureSource = null;
        excelFile = null;
    }

    @BeforeEach
    void createExcelFile() throws IOException {
        excelFile = Path.of(
                        System.getProperty("java.io.tmpdir"), getClass().getName() + "_" + UUID.randomUUID() + ".xlsx")
                .toFile();
        assumeTrue(excelFile.createNewFile(), () -> "Failed to create excel file" + excelFile.getAbsolutePath());
    }

    @BeforeEach
    void createRandomFeatures() throws IOException, SchemaException {
        SimpleFeatureType inputType =
                DataUtilities.createType(FT_NAME, "id:Integer,label:String,date:Date,location:Point:28992");
        MemoryDataStore dataStore = new MemoryDataStore(inputType);

        int[] xCoords = new Random().ints(randomFeatureCount, 155000, 165000).toArray();
        int[] yCoords = new Random().ints(randomFeatureCount, 463000, 473000).toArray();
        long minEpoch = LocalDate.of(2000, 1, 1).toEpochDay() * 86400L * 1000L;
        long maxEpoch = LocalDate.of(2025, 12, 31).toEpochDay() * 86400L * 1000L;
        Random random = new Random();
        SimpleFeatureBuilder fb = new SimpleFeatureBuilder(inputType);

        IntStream.range(0, randomFeatureCount).forEach(id -> {
            int x = xCoords[id];
            int y = yCoords[id];
            Date randomDate = new Date(minEpoch + (long) (random.nextDouble() * (maxEpoch - minEpoch)));
            fb.set("id", id);
            fb.set("date", randomDate);
            if (id == 1) {
                fb.set("label", "Null geometry Feature number " + id);
                fb.set("date", dateWithoutTime);
                fb.set("location", null);
            } else {
                fb.set("label", "Feature number " + id);
                GeometryFactory gf = new GeometryFactory();
                fb.set("location", gf.createPoint(new Coordinate(x, y)));
            }
            dataStore.addFeature(fb.buildFeature(String.valueOf(id)));
        });

        randomFeatureSource = dataStore.getFeatureSource(inputType.getName());
        assumeTrue(randomFeatureSource != null, "Failed to create random feature source");
    }

    @Test
    public void can_create_excel_file_from_feature_collection() throws IOException, SchemaException {
        Query q = new Query(randomFeatureSource.getName().toString());
        int featureCount = randomFeatureSource.getCount(q);
        assumeTrue(
                featureCount == randomFeatureCount,
                "Expected %s features in the input feature source, but got %s"
                        .formatted(randomFeatureCount, randomFeatureSource.getCount(q)));

        try (Transaction outputTransaction = new DefaultTransaction("tailormap-extract-output")) {
            ExcelDataStore ds = (ExcelDataStore) new ExcelDataStoreFactory()
                    .createDataStore(Map.of(
                            ExcelDataStoreFactory.FILE_PARAM.key,
                            excelFile,
                            ExcelDataStoreFactory.SHEET_PARAM.key,
                            FT_NAME));

            SimpleFeatureType subType =
                    DataUtilities.createSubType(randomFeatureSource.getSchema(), "label", "date", "location");
            ds.createSchema(subType);

            final AtomicInteger featuresAddedCounter = new AtomicInteger();
            if (ds.getFeatureSource() instanceof SimpleFeatureStore featureStore) {
                featureStore.setTransaction(outputTransaction);

                featureStore
                        .getSchema()
                        .getAttributeDescriptors()
                        .forEach(attr -> logger.finer("Output feature type has attribute " + attr.getLocalName()));

                featureStore.addFeatureListener(event -> {
                    if (event.getType().equals(FeatureEvent.Type.ADDED)) {
                        featuresAddedCounter.getAndIncrement();
                    }
                    if (featuresAddedCounter.get() % 50 == 0) {
                        logger.info("%s features added to Excel store. Progress: %s%%"
                                .formatted(featuresAddedCounter.get(), (int)
                                        ((featuresAddedCounter.doubleValue() / featureCount) * 100)));
                    }
                });
                List<FeatureId> ids = featureStore.addFeatures(randomFeatureSource.getFeatures(q));
                assertEquals(
                        randomFeatureCount, ids.size(), () -> "Expected %s features".formatted(randomFeatureCount));

                outputTransaction.commit();
            } else {
                fail("Output datastore is not a SimpleFeatureStore, cannot write features");
            }
            ds.dispose();
        }

        // open the Excel file and check that we have the expected content in the header
        try (InputStream inp = new FileInputStream(excelFile);
                Workbook wb = WorkbookFactory.create(inp)) {
            Sheet sheet = wb.getSheetAt(0);

            assertEquals(
                    randomFeatureCount + /*header row*/ 1,
                    sheet.getPhysicalNumberOfRows(),
                    "Expected %s rows in the Excel sheet".formatted(randomFeatureCount));

            assertAll(
                    "Check header row",
                    () -> assertEquals(FT_NAME, sheet.getSheetName(), "Expected sheet name to be " + FT_NAME),
                    () -> assertEquals(
                            3, sheet.getRow(0).getPhysicalNumberOfCells(), "Expected 3 columns in the header row"),
                    () -> assertEquals(
                            CellType.STRING,
                            sheet.getRow(0).getCell(0).getCellType(),
                            "Expected first cell in header to be a string"),
                    () -> assertEquals(
                            CellType.STRING,
                            sheet.getRow(0).getCell(1).getCellType(),
                            "Expected second cell in header to be a string"),
                    () -> assertEquals("label", sheet.getRow(0).getCell(0).getStringCellValue()),
                    () -> assertEquals("date", sheet.getRow(0).getCell(1).getStringCellValue()),
                    () -> assertEquals("location", sheet.getRow(0).getCell(2).getStringCellValue()));

            assertAll(
                    "Check first data row",
                    () -> assertTrue(
                            sheet.getRow(1).getCell(0).getStringCellValue().startsWith("Feature number ")),
                    () -> assertEquals(
                            CellType.NUMERIC,
                            sheet.getRow(1).getCell(1).getCellType(),
                            "Expected date cell to be NUMERIC (Excel date)"),
                    () -> assertThat(
                            "Expected value to be a valid date",
                            sheet.getRow(1).getCell(1).getDateCellValue(),
                            instanceOf(Date.class)),
                    () -> assertTrue(
                            sheet.getRow(1).getCell(2).getStringCellValue().startsWith("POINT")));

            assertAll(
                    "Check null geometry (2nd) row",
                    () -> assertTrue(
                            sheet.getRow(2).getCell(0).getStringCellValue().startsWith("Null geometry Feature number")),
                    () -> assumeTrue(
                            sheet.getPhysicalNumberOfRows() > 1,
                            "Expected more than 1 data row in the sheet to check the null geometry row"),
                    () -> assertEquals(
                            CellType.NUMERIC,
                            sheet.getRow(2).getCell(1).getCellType(),
                            "Expected date cell to be NUMERIC (Excel date)"),
                    () -> assertEquals(
                            dateWithoutTime,
                            sheet.getRow(2).getCell(1).getDateCellValue(),
                            "Expected date to be at midnight"),
                    () -> assertEquals(
                            CellType.BLANK, sheet.getRow(2).getCell(2).getCellType()));

            assertAll(
                    "Check last data row",
                    () -> assumeTrue(
                            sheet.getPhysicalNumberOfRows() > 2,
                            "Expected more than 2 rows in the sheet to check the last data row"),
                    () -> assertTrue(sheet.getRow(sheet.getPhysicalNumberOfRows() - 1)
                            .getCell(0)
                            .getStringCellValue()
                            .startsWith("Feature number ")),
                    () -> assertTrue(sheet.getRow(sheet.getPhysicalNumberOfRows() - 1)
                            .getCell(2)
                            .getStringCellValue()
                            .startsWith("POINT")));
        }
    }

    @Test
    public void oversized_text_value_is_truncated_with_comment_and_error_style() throws IOException, SchemaException {
        SimpleFeatureType oversizedType = DataUtilities.createType("oversized", "label:String");
        MemoryDataStore memStore = new MemoryDataStore(oversizedType);

        // Create a string value that exceeds Excel's max text length
        String oversizedValue = "x".repeat(SpreadsheetVersion.EXCEL2007.getMaxTextLength() + 1);

        SimpleFeatureBuilder fb = new SimpleFeatureBuilder(oversizedType);
        fb.set("label", oversizedValue);
        memStore.addFeature(fb.buildFeature("1"));

        SimpleFeatureSource source = memStore.getFeatureSource(oversizedType.getName());

        try (Transaction outputTransaction = new DefaultTransaction("test-oversized")) {
            ExcelDataStore ds = (ExcelDataStore) new ExcelDataStoreFactory()
                    .createDataStore(Map.of(
                            ExcelDataStoreFactory.FILE_PARAM.key,
                            excelFile,
                            ExcelDataStoreFactory.SHEET_PARAM.key,
                            "oversized"));

            ds.createSchema(oversizedType);

            if (ds.getFeatureSource() instanceof SimpleFeatureStore featureStore) {
                featureStore.setTransaction(outputTransaction);
                featureStore.addFeatures(source.getFeatures());
                outputTransaction.commit();
            } else {
                fail("Output datastore is not a SimpleFeatureStore");
            }
            ds.dispose();
        } finally {
            memStore.dispose();
        }

        try (InputStream inp = new FileInputStream(excelFile);
                Workbook wb = WorkbookFactory.create(inp)) {
            Sheet sheet = wb.getSheetAt(0);
            Row dataRow = sheet.getRow(1);
            Cell cell = dataRow.getCell(0);

            assertAll(
                    "Check oversized cell value behavior",
                    () -> {
                        String cellValue = cell.getStringCellValue();
                        assertTrue(cellValue.endsWith(" ..."), "Expected truncated value to end with ' ...'");
                        assertEquals(
                                SpreadsheetVersion.EXCEL2007.getMaxTextLength(),
                                cellValue.length(),
                                "Expected truncated value to have max text length");
                    },
                    () -> assertNotNull(cell.getCellComment(), "Expected a cell comment for oversized value"),
                    () -> assertEquals(
                            "Source value greater than maximum allowed for Excel",
                            cell.getCellComment().getString().getString(),
                            "Expected specific comment text"),
                    () -> assertEquals(
                            IndexedColors.RED.getIndex(),
                            wb.getFontAt(cell.getCellStyle().getFontIndex()).getColor(),
                            "Expected red font color for error style"));
        }
    }
}
