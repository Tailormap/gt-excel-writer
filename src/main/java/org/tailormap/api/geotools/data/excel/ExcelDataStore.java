/*
 * Copyright (C) 2026 B3Partners B.V.
 *
 * SPDX-License-Identifier: LGPL-2.1-or-later
 */
package org.tailormap.api.geotools.data.excel;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Collections;
import java.util.List;
import java.util.logging.Level;
import java.util.logging.Logger;
import org.apache.poi.ss.SpreadsheetVersion;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.geotools.api.data.FeatureReader;
import org.geotools.api.data.FeatureWriter;
import org.geotools.api.data.FileDataStore;
import org.geotools.api.data.Query;
import org.geotools.api.data.SimpleFeatureSource;
import org.geotools.api.data.Transaction;
import org.geotools.api.feature.simple.SimpleFeature;
import org.geotools.api.feature.simple.SimpleFeatureType;
import org.geotools.api.feature.type.Name;
import org.geotools.api.filter.Filter;
import org.geotools.data.store.ContentDataStore;
import org.geotools.data.store.ContentEntry;
import org.geotools.data.store.ContentFeatureSource;
import org.geotools.feature.NameImpl;
import org.geotools.util.logging.Logging;

public class ExcelDataStore extends ContentDataStore implements FileDataStore {
  private SimpleFeatureType schema;
  private NameImpl typeName;
  private final SXSSFWorkbook workbook;
  private final File excelFile;
  private boolean workbookDisposed = false;
  private boolean enableCellAutoSizing;
  private static final Logger logger = Logging.getLogger(ExcelDataStore.class);

  /**
   * Create an ExcelDataStore for writing to the specified Excel file.
   *
   * @param typeName the FeatureType name
   * @param excelFile the file to write the Excel data to. The file will must be created/exist.
   */
  public ExcelDataStore(NameImpl typeName, File excelFile) {
    this(typeName, excelFile, false);
  }

  /**
   * Create an ExcelDataStore for writing to the specified Excel file.
   *
   * @param typeName the FeatureType name
   * @param excelFile the file to write the Excel data to. The file will must be created/exist.
   * @param enableCellAutoSizing whether to enable auto-sizing of cells when writing all features. This can
   *     significantly impact performance for writing large datasets. Defaults to false.
   */
  public ExcelDataStore(NameImpl typeName, File excelFile, boolean enableCellAutoSizing) {
    this.typeName = typeName;
    this.excelFile = excelFile;
    this.enableCellAutoSizing = enableCellAutoSizing;
    // create a new workbook with a small window size for streaming
    this.workbook = new SXSSFWorkbook(1);
    // TODO / EVALUATE avoid performance penalty of compressing the temporary files,
    // this.workbook.setCompressTempFiles(true);

    // setting this option will make excel and old versions of open/libre office and google docs break,
    // now it is only broken for current libreoffice
    // see https://bugs.documentfoundation.org/show_bug.cgi?id=163384
    // this.workbook.setZip64Mode(Zip64Mode.AlwaysWithCompatibility);
  }

  public SXSSFWorkbook getWorkbook() {
    return workbook;
  }

  public static int getMaxRows() {
    return SpreadsheetVersion.EXCEL2007.getMaxRows();
  }

  public static int getMaxColumns() {
    return SpreadsheetVersion.EXCEL2007.getMaxColumns();
  }

  @Override
  public void dispose() {
    if (!workbookDisposed) {
      try (workbook;
          FileOutputStream out = new FileOutputStream(excelFile)) {
        workbook.write(out);
        workbook.close();
        out.flush();
      } catch (IOException e) {
        logger.log(Level.SEVERE, "Error writing Excel workbook to file", e);
      } finally {
        workbookDisposed = true;
      }
    }
    super.dispose();
  }

  @Override
  public SimpleFeatureSource getFeatureSource() throws IOException {
    if (typeName == null) {
      createTypeNames();
    }
    return super.getFeatureSource(typeName);
  }

  @Override
  public FeatureWriter<SimpleFeatureType, SimpleFeature> getFeatureWriter(Filter filter, Transaction transaction)
      throws IOException {
    // not used in our write-only implementation scenario
    return super.getFeatureWriter(typeName.getLocalPart(), filter, transaction);
  }

  @Override
  public FeatureWriter<SimpleFeatureType, SimpleFeature> getFeatureWriter(Transaction transaction)
      throws IOException {
    // not used in our write-only implementation scenario
    return super.getFeatureWriter(typeName.getLocalPart(), transaction);
  }

  @Override
  public FeatureWriter<SimpleFeatureType, SimpleFeature> getFeatureWriterAppend(Transaction transaction)
      throws IOException {
    // not used in our write-only implementation scenario
    return super.getFeatureWriterAppend(typeName.getLocalPart(), transaction);
  }

  @Override
  protected ContentFeatureSource createFeatureSource(ContentEntry entry) throws IOException {
    if (excelFile != null && excelFile.canWrite()) {
      return new ExcelFeatureStore(entry, Query.ALL, this.enableCellAutoSizing);
    } else {
      logger.severe("Cannot write Excel File");
      // TODO not used?
      return new ExcelFeatureSource(entry, Query.ALL);
    }
  }

  @Override
  public SimpleFeatureType getSchema() throws IOException {
    if (schema == null) {
      schema = getSchema(typeName);
    }
    return schema;
  }

  public Name getTypeName() {
    if (namespaceURI != null) {
      return new NameImpl(namespaceURI, typeName.getLocalPart());
    } else {
      return typeName;
    }
  }

  @Override
  protected List<Name> createTypeNames() {
    return Collections.singletonList(getTypeName());
  }

  @Override
  public void updateSchema(SimpleFeatureType featureType) {
    schema = featureType;
  }

  @Override
  public void createSchema(SimpleFeatureType featureType) {
    schema = featureType;
    typeName = (NameImpl) schema.getName();
  }

  @Override
  public FeatureReader<SimpleFeatureType, SimpleFeature> getFeatureReader() {
    throw new UnsupportedOperationException("ExcelDataStore is write-only, cannot get reader");
  }

  public boolean isEnableCellAutoSizing() {
    return enableCellAutoSizing;
  }

  public void setEnableCellAutoSizing(boolean enableCellAutoSizing) {
    this.enableCellAutoSizing = enableCellAutoSizing;
  }
}
