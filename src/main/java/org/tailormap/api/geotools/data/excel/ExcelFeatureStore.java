/*
 * Copyright (C) 2026 B3Partners B.V.
 *
 * SPDX-License-Identifier: LGPL-2.1-or-later
 */
package org.tailormap.api.geotools.data.excel;

import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.util.logging.Logger;
import org.geotools.api.data.FeatureReader;
import org.geotools.api.data.FeatureWriter;
import org.geotools.api.data.Query;
import org.geotools.api.data.Transaction;
import org.geotools.api.feature.simple.SimpleFeature;
import org.geotools.api.feature.simple.SimpleFeatureType;
import org.geotools.api.feature.type.AttributeDescriptor;
import org.geotools.api.filter.identity.FeatureId;
import org.geotools.data.store.ContentEntry;
import org.geotools.data.store.ContentFeatureStore;
import org.geotools.feature.FeatureCollection;
import org.geotools.feature.FeatureIterator;
import org.geotools.geometry.jts.ReferencedEnvelope;
import org.geotools.util.logging.Logging;

/**
 * Feature store implementation for Excel-based data backed by an {@link ExcelDataStore}.
 *
 * <p>This class provides read and write access to features stored in Excel sheets by delegating to
 * {@link ExcelFeatureReader} and {@link ExcelFeatureWriter}. It intentionally implements only the operations required
 * for the current use-cases; methods such as {@link #getBoundsInternal(Query)} and {@link #getCountInternal(Query)} are
 * not implemented and will return conservative defaults.
 *
 * <p>Instances are created with an option to enable cell auto-sizing when writing features. Enabling auto-sizing can
 * improve output appearance but may significantly impact performance for large datasets; the default is disabled.
 */
public class ExcelFeatureStore extends ContentFeatureStore {
    private final boolean enableCellAutoSizing;
    private static final Logger logger = Logging.getLogger(ExcelFeatureStore.class);

    /**
     * Creates the content feature store, with cell auto-sizing disabled
     *
     * @param entry The content entry this feature store belongs to.
     * @param query The defining query.
     */
    public ExcelFeatureStore(ContentEntry entry, Query query) {
        super(entry, query);
        this.enableCellAutoSizing = false;
    }

    /**
     * Creates the content feature store.
     *
     * @param entry The content entry this feature store belongs to
     * @param query The defining query
     * @param enableCellAutoSizing whether to enable auto-sizing of cells when writing all features. This can
     *     significantly impact performance when writing large datasets. Defaults to false.
     */
    public ExcelFeatureStore(ContentEntry entry, Query query, boolean enableCellAutoSizing) {
        super(entry, query);
        this.enableCellAutoSizing = enableCellAutoSizing;
    }

    @Override
    public List<FeatureId> addFeatures(FeatureCollection<SimpleFeatureType, SimpleFeature> featureCollection)
            throws IOException {

        List<FeatureId> writtenFeatureIds = new ArrayList<>();
        try (FeatureWriter<SimpleFeatureType, SimpleFeature> writer = getWriterInternal(Query.ALL, WRITER_ADD);
                FeatureIterator<SimpleFeature> iterator = featureCollection.features()) {
            while (iterator.hasNext()) {
                SimpleFeature sourceFeature = iterator.next();
                SimpleFeature targetFeature = writer.next();
                logger.finest("Source feature " + sourceFeature.getAttributes());
                for (AttributeDescriptor att : targetFeature.getFeatureType().getAttributeDescriptors()) {
                    targetFeature.setAttribute(att.getName(), sourceFeature.getAttribute(att.getName()));
                }
                logger.finest("Target feature (updated) " + targetFeature.getAttributes());
                writer.write();
                if (sourceFeature.getIdentifier() != null) {
                    writtenFeatureIds.add(sourceFeature.getIdentifier());
                }
            }
        }

        return writtenFeatureIds;
    }

    @Override
    protected FeatureWriter<SimpleFeatureType, SimpleFeature> getWriterInternal(Query query, int flags)
            throws IOException {
        return new ExcelFeatureWriter(entry, query, enableCellAutoSizing);
    }

    @Override
    protected FeatureReader<SimpleFeatureType, SimpleFeature> getReaderInternal(Query query) throws IOException {
        return new ExcelFeatureReader(entry, query);
    }

    @Override
    public void setTransaction(Transaction transaction) {
        super.setTransaction(transaction);
    }

    @Override
    protected ReferencedEnvelope getBoundsInternal(Query query) {
        logger.severe("getBoundsInternal is not implemented for ExcelFeatureStore, returning null");
        return null;
    }

    @Override
    protected int getCountInternal(Query query) {
        logger.severe("getCount is not implemented for ExcelFeatureStore, returning -1");
        return -1;
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
