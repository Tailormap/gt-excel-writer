/*
 * Copyright (C) 2026 B3Partners B.V.
 *
 * SPDX-License-Identifier: LGPL-2.1-or-later
 */
package org.tailormap.api.geotools.data.excel;

import java.io.IOException;
import java.util.logging.Logger;
import org.geotools.api.data.FeatureReader;
import org.geotools.api.data.Query;
import org.geotools.api.feature.simple.SimpleFeature;
import org.geotools.api.feature.simple.SimpleFeatureType;
import org.geotools.data.store.ContentEntry;
import org.geotools.data.store.ContentFeatureSource;
import org.geotools.geometry.jts.ReferencedEnvelope;
import org.geotools.util.logging.Logging;

public class ExcelFeatureSource extends ContentFeatureSource {
    private static final Logger logger = Logging.getLogger(ExcelFeatureSource.class);

    public ExcelFeatureSource(ContentEntry entry, Query query) {
        super(entry, query);
        transaction = getState().getTransaction();
    }

    public ExcelFeatureSource(ExcelDataStore datastore) {
        this(datastore, Query.ALL);
    }

    public ExcelFeatureSource(ExcelDataStore datastore, Query query) {
        this(new ContentEntry(datastore, datastore.getTypeName()), query);
    }

    public ExcelFeatureSource(ContentEntry entry) throws IOException {
        this(entry, Query.ALL);
    }

    @Override
    protected ReferencedEnvelope getBoundsInternal(Query query) throws IOException {
        logger.warning("getBoundsInternal is not implemented for ExcelFeatureSource, returning null");
        return null;
    }

    @Override
    protected int getCountInternal(Query query) throws IOException {
        logger.warning("getCount is not implemented for ExcelFeatureSource, returning -1");
        return -1;
    }

    @Override
    protected FeatureReader<SimpleFeatureType, SimpleFeature> getReaderInternal(Query query) throws IOException {
        throw new UnsupportedOperationException("ExcelDataStore is write-only, cannot get reader");
    }

    @Override
    public ExcelDataStore getDataStore() {
        return (ExcelDataStore) entry.getDataStore();
    }

    @Override
    protected SimpleFeatureType buildFeatureType() throws IOException {
        return getDataStore().getSchema();
    }
}
