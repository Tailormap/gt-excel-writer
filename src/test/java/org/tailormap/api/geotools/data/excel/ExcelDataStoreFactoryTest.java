/*
 * Copyright (C) 2026 B3Partners B.V.
 *
 * SPDX-License-Identifier: LGPL-2.1-or-later
 */
package org.tailormap.api.geotools.data.excel;

import static org.junit.jupiter.api.Assertions.assertAll;
import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;
import static org.junit.jupiter.api.Assertions.fail;
import static org.junit.jupiter.api.Assumptions.assumeTrue;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Path;
import java.util.List;
import java.util.Map;
import java.util.Random;
import java.util.UUID;
import java.util.concurrent.atomic.AtomicInteger;
import java.util.logging.Logger;
import java.util.stream.IntStream;
import org.apache.poi.ss.usermodel.CellType;
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
import org.geotools.util.logging.Logging;
import org.junit.jupiter.api.AfterEach;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;

class ExcelDataStoreFactoryTest {
    private static final Logger logger = Logging.getLogger(ExcelDataStoreFactoryTest.class);
    private static final String FT_NAME = "random";
    private static final int randomFeatureCount = 104;
    private SimpleFeatureSource randomFeatureSource;
    private File excelFile;

    @AfterEach
    void tearDown() {
        randomFeatureSource.getDataStore().dispose();
        // the output file is left in the tmp directory for inspection
        // excelFile.deleteOnExit();
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
        SimpleFeatureType inputType = DataUtilities.createType(FT_NAME, "id:Integer,label:String,location:Point:28992");
        MemoryDataStore dataStore = new MemoryDataStore(inputType);

        int[] xCoords = new Random().ints(randomFeatureCount, 155000, 165000).toArray();
        int[] yCoords = new Random().ints(randomFeatureCount, 463000, 473000).toArray();

        IntStream.range(0, randomFeatureCount).forEach(id -> {
            int x = xCoords[id];
            int y = yCoords[id];

            if (id == 1) {
                dataStore.addFeature(
                        DataUtilities.createFeature(inputType, id + "|Null geometry Feature number " + id + "|null"));
            } else {
                dataStore.addFeature(DataUtilities.createFeature(
                        inputType, id + "|Feature number " + id + "|POINT(" + x + " " + y + ")"));
            }
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
                    DataUtilities.createSubType(randomFeatureSource.getSchema(), "label", "location");
            ds.createSchema(subType);

            final AtomicInteger featuresAddedCounter = new AtomicInteger();
            if (ds.getFeatureSource() instanceof SimpleFeatureStore featureStore) {
                featureStore.setTransaction(outputTransaction);

                featureStore
                        .getSchema()
                        .getAttributeDescriptors()
                        .forEach(attr -> logger.fine("Output feature type has attribute " + attr.getLocalName()));

                featureStore.addFeatureListener(event -> {
                    if (event.getType().equals(FeatureEvent.Type.ADDED)) {
                        featuresAddedCounter.getAndIncrement();
                    }
                    if (featuresAddedCounter.get() % 50 == 0) {
                        logger.info("%s features added. Progress: %s%%".formatted(featuresAddedCounter.get(), (int)
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

        // open the Excel file and check that we have the expected content
        try (InputStream inp = new FileInputStream(excelFile);
                Workbook wb = WorkbookFactory.create(inp)) {

            Sheet sheet = wb.getSheetAt(0);

            assertAll(
                    "Check header and first data row",
                    () -> assertEquals(FT_NAME, sheet.getSheetName(), "Expected sheet name to be " + FT_NAME),
                    () -> assertEquals(
                            2, sheet.getRow(0).getPhysicalNumberOfCells(), "Expected 2 columns in the header row"),
                    () -> assertEquals(
                            CellType.STRING,
                            sheet.getRow(0).getCell(0).getCellType(),
                            "Expected first cell in header to be a string"),
                    () -> assertEquals(
                            CellType.STRING,
                            sheet.getRow(0).getCell(1).getCellType(),
                            "Expected second cell in header to be a string"),
                    () -> assertEquals("label", sheet.getRow(0).getCell(0).getStringCellValue()),
                    () -> assertEquals("location", sheet.getRow(0).getCell(1).getStringCellValue()),
                    () -> assertTrue(
                            sheet.getRow(1).getCell(0).getStringCellValue().startsWith("Feature number ")),
                    () -> assertTrue(
                            sheet.getRow(1).getCell(1).getStringCellValue().startsWith("POINT")));

            assertEquals(
                    randomFeatureCount + /*header row*/ 1,
                    sheet.getPhysicalNumberOfRows(),
                    "Expected %s rows in the Excel sheet".formatted(randomFeatureCount));

            assertAll(
                    "Check null geometry row",
                    () -> assumeTrue(
                            sheet.getPhysicalNumberOfRows() > 1,
                            "Expected more than 1 data rows in the sheet to check the null geometry row"),
                    () -> assertTrue(sheet.getRow(2)
                            .getCell(0)
                            .getStringCellValue()
                            .startsWith("Null geometry Feature number ")),
                    () -> assertEquals(
                            CellType.BLANK, sheet.getRow(2).getCell(1).getCellType()));
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
                            .getCell(1)
                            .getStringCellValue()
                            .startsWith("POINT")));
        }
    }
}
