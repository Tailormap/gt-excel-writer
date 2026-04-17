/*
 * Copyright (C) 2026 B3Partners B.V.
 *
 * SPDX-License-Identifier: LGPL-2.1-or-later
 */
package org.tailormap.api.geotools.data.excel;

import java.io.File;
import java.io.IOException;
import java.net.URL;
import java.util.Map;
import java.util.Objects;
import org.apache.commons.io.FilenameUtils;
import org.geotools.api.data.DataStore;
import org.geotools.api.data.FileDataStore;
import org.geotools.api.data.FileDataStoreFactorySpi;
import org.geotools.feature.NameImpl;
import org.geotools.util.URLs;

public class ExcelDataStoreFactory implements FileDataStoreFactorySpi {
    public static final Param FILE_PARAM = new Param("file", File.class, "Excel file to write the data", true);
    public static final Param SHEET_PARAM =
            new Param("sheet", String.class, "Name of the sheet to write the data, eg. the feature type name", false);
    private static final String FILE_TYPE = "xlsx";
    private static final String[] EXTENSIONS = {"." + FILE_TYPE};
    private static final Param[] PARAMETERS_INFO = {FILE_PARAM};

    @Override
    public FileDataStore createDataStore(URL url) throws IOException {
        File file = URLs.urlToFile(url);
        return (ExcelDataStore) this.createDataStore(Map.of(FILE_PARAM.key, file));
    }

    @Override
    public DataStore createDataStore(Map<String, ?> params) throws IOException {
        File file = (File) FILE_PARAM.lookUp(params);
        if (file == null) {
            throw new IOException("No file parameter provided");
        }
        String sheetName = (String) SHEET_PARAM.lookUp(params);
        if (sheetName == null) {
            sheetName = FilenameUtils.getBaseName(file.getPath());
        }
        return new ExcelDataStore(new NameImpl(sheetName), file);
    }

    @Override
    public DataStore createNewDataStore(Map<String, ?> params) throws IOException {
        return this.createDataStore(params);
    }

    @Override
    public String getTypeName(URL url) {
        throw new UnsupportedOperationException("Not implemented.");
    }

    /**
     * @param url URL to a real file (may not be local)
     * @return true if the file extension is supported, false otherwise
     */
    @Override
    public boolean canProcess(URL url) {
        return FILE_TYPE.equalsIgnoreCase(FilenameUtils.getExtension(
                Objects.requireNonNull(URLs.urlToFile(url)).toString()));
    }

    @Override
    public String[] getFileExtensions() {
        return EXTENSIONS;
    }

    @Override
    public String getDisplayName() {
        return "Excel DataStore (write only)";
    }

    @Override
    public String getDescription() {
        return "A Datastore backed by an Excel Workbook";
    }

    @Override
    public Param[] getParametersInfo() {
        return PARAMETERS_INFO;
    }

    @Override
    public boolean isAvailable() {
        try {
            String ignored = ExcelDataStore.class.getName();
        } catch (Exception e) {
            return false;
        }
        return true;
    }
}
